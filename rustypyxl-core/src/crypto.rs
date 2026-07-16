//! Reading and writing password-protected (ECMA-376 encrypted) workbooks.
//!
//! An encrypted xlsx is not a ZIP: it is an OLE2 / Compound File Binary
//! container holding an `EncryptionInfo` stream (the scheme) and an
//! `EncryptedPackage` stream (the AES-encrypted ZIP). Reading decrypts that
//! package into the plaintext ZIP bytes (the normal load path); writing
//! (the `encrypt` feature) does the reverse, producing the container from a
//! saved ZIP -- validated against msoffcrypto-tool.
//!
//! Only **agile encryption** (Office 2010+, EncryptionInfo version 4.4) is
//! supported -- effectively all modern files. The AES block cipher and SHA
//! hashes come from the RustCrypto crates already in the build (zip pulls them
//! for AES-encrypted ZIP entries); only the compound-file reader is new. The
//! CBC chaining is done here rather than via another crate, since it is a plain
//! block-XOR wrapper, not a cryptographic primitive.

use crate::error::{Result, RustypyxlError};
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha384, Sha512};

/// The OLE2/CFB magic that marks an encrypted OOXML container (vs a ZIP's "PK").
const CFB_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

// Fixed block keys from ECMA-376 for deriving per-purpose keys.
const BLOCK_KEY_VERIFIER_INPUT: [u8; 8] = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
const BLOCK_KEY_VERIFIER_VALUE: [u8; 8] = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
const BLOCK_KEY_ENCRYPTED_KEY: [u8; 8] = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];

/// Whether the bytes look like an encrypted OOXML container (a CFB file).
pub fn is_encrypted(data: &[u8]) -> bool {
    data.len() >= 8 && data[..8] == CFB_MAGIC
}

/// Decrypt an encrypted workbook into the plaintext ZIP bytes it wraps.
pub fn decrypt(data: &[u8], password: &str) -> Result<Vec<u8>> {
    let container = cfb::Container::open(data)?;
    let info = container.stream("EncryptionInfo")?;
    let package = container.stream("EncryptedPackage")?;

    if info.len() < 8 {
        return Err(RustypyxlError::InvalidFormat(
            "truncated EncryptionInfo".into(),
        ));
    }
    let major = u16::from_le_bytes([info[0], info[1]]);
    let minor = u16::from_le_bytes([info[2], info[3]]);
    if (major, minor) == (4, 4) {
        decrypt_agile(&info[8..], &package, password)
    } else {
        Err(RustypyxlError::InvalidFormat(format!(
            "unsupported encryption version {major}.{minor}; only agile (4.4) is supported"
        )))
    }
}

/// A minimal read-only OLE2 / Compound File Binary reader -- just enough to pull
/// named streams out of an encrypted workbook's container. Only what the format
/// needs here is implemented (FAT + mini-FAT chains, the directory, and stream
/// reads); it is not a general CFB library.
mod cfb {
    use crate::error::{Result, RustypyxlError};

    const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
    const FREESECT: u32 = 0xFFFF_FFFF;

    pub struct Container {
        data: Vec<u8>,
        sector_size: usize,
        mini_sector_size: usize,
        mini_cutoff: u32,
        fat: Vec<u32>,
        mini_fat: Vec<u32>,
        /// (name, object_type, start_sector, size) per directory entry.
        dir: Vec<(String, u8, u32, u64)>,
        mini_stream: Vec<u8>,
    }

    fn u16le(d: &[u8], off: usize) -> u16 {
        u16::from_le_bytes([d[off], d[off + 1]])
    }
    fn u32le(d: &[u8], off: usize) -> u32 {
        u32::from_le_bytes([d[off], d[off + 1], d[off + 2], d[off + 3]])
    }
    fn u64le(d: &[u8], off: usize) -> u64 {
        u64::from_le_bytes(d[off..off + 8].try_into().unwrap())
    }

    impl Container {
        pub fn open(data: &[u8]) -> Result<Self> {
            if data.len() < 512 || data[..8] != super::CFB_MAGIC {
                return Err(RustypyxlError::InvalidFormat("not a compound file".into()));
            }
            let sector_size = 1usize << u16le(data, 30);
            let mini_sector_size = 1usize << u16le(data, 32);
            let num_fat_sectors = u32le(data, 44);
            let first_dir_sector = u32le(data, 48);
            let mini_cutoff = u32le(data, 56);
            let first_mini_fat = u32le(data, 60);
            let num_mini_fat = u32le(data, 64);

            let mut c = Container {
                data: data.to_vec(),
                sector_size,
                mini_sector_size,
                mini_cutoff,
                fat: Vec::new(),
                mini_fat: Vec::new(),
                dir: Vec::new(),
                mini_stream: Vec::new(),
            };

            // Build the FAT from the DIFAT (first 109 entries live in the header;
            // small files never spill past them).
            let mut fat_sectors = Vec::new();
            for i in 0..num_fat_sectors.min(109) as usize {
                let s = u32le(&c.data, 76 + i * 4);
                if s != FREESECT && s != ENDOFCHAIN {
                    fat_sectors.push(s);
                }
            }
            for s in fat_sectors {
                let bytes = c.sector_bytes(s)?;
                for k in 0..sector_size / 4 {
                    c.fat.push(u32le(&bytes, k * 4));
                }
            }

            // Mini-FAT chain.
            if first_mini_fat != ENDOFCHAIN && num_mini_fat > 0 {
                let bytes = c.read_fat_chain(first_mini_fat)?;
                for k in 0..bytes.len() / 4 {
                    c.mini_fat.push(u32le(&bytes, k * 4));
                }
            }

            // Directory entries.
            let dir_bytes = c.read_fat_chain(first_dir_sector)?;
            for entry in dir_bytes.chunks(128) {
                if entry.len() < 128 {
                    break;
                }
                let name_len = u16le(entry, 64) as usize;
                let obj_type = entry[66];
                if obj_type == 0 || name_len < 2 {
                    continue;
                }
                let name: String = (0..(name_len / 2 - 1))
                    .map(|i| u16le(entry, i * 2))
                    .filter_map(|u| char::from_u32(u as u32))
                    .collect();
                let start = u32le(entry, 116);
                let size = u64le(entry, 120);
                c.dir.push((name, obj_type, start, size));
            }

            // The root entry (type 5) owns the mini stream.
            if let Some((start, size)) = c
                .dir
                .iter()
                .find(|(_, t, _, _)| *t == 5)
                .map(|(_, _, s, sz)| (*s, *sz))
            {
                let mut mini = c.read_fat_chain(start)?;
                mini.truncate(size as usize);
                c.mini_stream = mini;
            }
            Ok(c)
        }

        /// The raw bytes of one big FAT sector.
        fn sector_bytes(&self, sector: u32) -> Result<Vec<u8>> {
            let start = (sector as usize + 1) * self.sector_size;
            let end = start + self.sector_size;
            self.data
                .get(start..end)
                .map(|s| s.to_vec())
                .ok_or_else(|| RustypyxlError::InvalidFormat("CFB sector out of range".into()))
        }

        /// Follow a FAT chain from `start`, concatenating its big sectors.
        fn read_fat_chain(&self, start: u32) -> Result<Vec<u8>> {
            let mut out = Vec::new();
            let mut cur = start;
            let mut guard = 0;
            while cur != ENDOFCHAIN && cur != FREESECT {
                out.extend_from_slice(&self.sector_bytes(cur)?);
                cur = *self
                    .fat
                    .get(cur as usize)
                    .ok_or_else(|| RustypyxlError::InvalidFormat("bad FAT chain".into()))?;
                guard += 1;
                if guard > self.fat.len() + 1 {
                    return Err(RustypyxlError::InvalidFormat("cyclic FAT chain".into()));
                }
            }
            Ok(out)
        }

        /// Follow a mini-FAT chain from `start`, concatenating its mini sectors
        /// out of the mini stream.
        fn read_mini_chain(&self, start: u32, size: u64) -> Result<Vec<u8>> {
            let mut out = Vec::new();
            let mut cur = start;
            let mut guard = 0;
            while cur != ENDOFCHAIN && cur != FREESECT {
                let off = cur as usize * self.mini_sector_size;
                let end = off + self.mini_sector_size;
                out.extend_from_slice(
                    self.mini_stream
                        .get(off..end)
                        .ok_or_else(|| RustypyxlError::InvalidFormat("mini sector oob".into()))?,
                );
                cur = *self
                    .mini_fat
                    .get(cur as usize)
                    .ok_or_else(|| RustypyxlError::InvalidFormat("bad mini-FAT chain".into()))?;
                guard += 1;
                if guard > self.mini_fat.len() + 1 {
                    return Err(RustypyxlError::InvalidFormat(
                        "cyclic mini-FAT chain".into(),
                    ));
                }
            }
            out.truncate(size as usize);
            Ok(out)
        }

        /// Read a named stream. Small streams (< mini cutoff) come from the
        /// mini-FAT; larger ones from the regular FAT.
        pub fn stream(&self, name: &str) -> Result<Vec<u8>> {
            let (start, size) = self
                .dir
                .iter()
                .find(|(n, t, _, _)| *t == 2 && n == name)
                .map(|(_, _, s, sz)| (*s, *sz))
                .ok_or_else(|| RustypyxlError::InvalidFormat(format!("missing {name} stream")))?;
            let want = size as usize;
            // A sub-cutoff stream normally lives in the mini stream, but some
            // writers (msoffcrypto) place it in the regular FAT instead; if the
            // mini-FAT read comes up short, fall back to the regular FAT.
            if size < self.mini_cutoff as u64 {
                if let Ok(mini) = self.read_mini_chain(start, size) {
                    if mini.len() >= want {
                        return Ok(mini);
                    }
                }
            }
            let mut b = self.read_fat_chain(start)?;
            b.truncate(want);
            Ok(b)
        }
    }

    // ---------- writing ----------

    #[cfg(feature = "encrypt")]
    const FATSECT: u32 = 0xFFFF_FFFD;
    #[cfg(feature = "encrypt")]
    const NOSTREAM: u32 = 0xFFFF_FFFF;
    #[cfg(feature = "encrypt")]
    const SECTOR: usize = 512;
    #[cfg(feature = "encrypt")]
    const MINI: usize = 64;
    #[cfg(feature = "encrypt")]
    const CUTOFF: usize = 4096;

    /// Directory-entry name order: shorter names first, then case-insensitively.
    #[cfg(feature = "encrypt")]
    fn cfb_name_order(a: &str, b: &str) -> std::cmp::Ordering {
        let (la, lb) = (a.encode_utf16().count(), b.encode_utf16().count());
        la.cmp(&lb)
            .then_with(|| a.to_uppercase().cmp(&b.to_uppercase()))
    }

    /// Round `n` up to a multiple of `unit`.
    #[cfg(feature = "encrypt")]
    fn round_up(n: usize, unit: usize) -> usize {
        n.div_ceil(unit) * unit
    }

    /// Write a compound file containing the given named streams. Streams smaller
    /// than the mini cutoff go in the mini stream; larger ones in the regular
    /// FAT. Enough for an encrypted workbook (root plus two streams).
    #[cfg(feature = "encrypt")]
    pub fn write_container(streams: &[(&str, &[u8])]) -> super::Result<Vec<u8>> {
        use super::RustypyxlError;

        // Order entries the way readers expect (root is entry 0).
        let mut ordered: Vec<(&str, &[u8])> = streams.to_vec();
        ordered.sort_by(|a, b| cfb_name_order(a.0, b.0));

        // Partition into mini (small) and big streams; build the mini stream and
        // its mini-FAT.
        let mut mini_stream = Vec::new();
        let mut mini_fat: Vec<u32> = Vec::new();
        // entry index (1-based) -> (start sector-or-minisector, is_mini)
        let mut placement: Vec<(u32, bool)> = Vec::with_capacity(ordered.len());
        for (_, data) in &ordered {
            if data.len() < CUTOFF {
                let start_mini = (mini_stream.len() / MINI) as u32;
                let nsec = data.len().div_ceil(MINI).max(1);
                for k in 0..nsec {
                    mini_fat.push(if k + 1 < nsec {
                        start_mini + k as u32 + 1
                    } else {
                        ENDOFCHAIN
                    });
                }
                let mut padded = data.to_vec();
                padded.resize(round_up(data.len().max(1), MINI), 0);
                mini_stream.extend_from_slice(&padded);
                placement.push((start_mini, true));
            } else {
                placement.push((0, false)); // filled in once big-stream sectors are laid out
            }
        }

        // Count regular sectors.
        let dir_entries = 1 + ordered.len();
        let dir_sectors = (dir_entries * 128).div_ceil(SECTOR);
        let minifat_sectors = if mini_fat.is_empty() {
            0
        } else {
            (mini_fat.len() * 4).div_ceil(SECTOR)
        };
        let ministream_sectors = mini_stream.len().div_ceil(SECTOR);
        let big_sectors: Vec<usize> = ordered
            .iter()
            .map(|(_, d)| {
                if d.len() >= CUTOFF {
                    d.len().div_ceil(SECTOR)
                } else {
                    0
                }
            })
            .collect();
        let non_fat: usize =
            dir_sectors + minifat_sectors + ministream_sectors + big_sectors.iter().sum::<usize>();
        let mut num_fat = 1usize;
        while non_fat + num_fat > num_fat * (SECTOR / 4) {
            num_fat += 1;
        }

        // Assign sector numbers: FAT, directory, mini-FAT, mini stream, big streams.
        let dir_start = num_fat as u32;
        let minifat_start = dir_start + dir_sectors as u32;
        let ministream_start = minifat_start + minifat_sectors as u32;
        let mut cursor = ministream_start + ministream_sectors as u32;
        for (i, (_, d)) in ordered.iter().enumerate() {
            if d.len() >= CUTOFF {
                placement[i] = (cursor, false);
                cursor += big_sectors[i] as u32;
            }
        }
        let total = cursor as usize;

        // Build the regular FAT.
        let mut fat = vec![FREESECT; num_fat * (SECTOR / 4)];
        let chain = |fat: &mut Vec<u32>, start: u32, count: usize| {
            for k in 0..count {
                let s = start as usize + k;
                fat[s] = if k + 1 < count {
                    start + k as u32 + 1
                } else {
                    ENDOFCHAIN
                };
            }
        };
        for slot in fat.iter_mut().take(num_fat) {
            *slot = FATSECT;
        }
        chain(&mut fat, dir_start, dir_sectors);
        if minifat_sectors > 0 {
            chain(&mut fat, minifat_start, minifat_sectors);
        }
        if ministream_sectors > 0 {
            chain(&mut fat, ministream_start, ministream_sectors);
        }
        for (i, (_, d)) in ordered.iter().enumerate() {
            if d.len() >= CUTOFF {
                chain(&mut fat, placement[i].0, big_sectors[i]);
            }
        }
        if total > fat.len() {
            return Err(RustypyxlError::Custom("CFB layout overflow".into()));
        }

        // Directory entries.
        let mut dir = Vec::new();
        let write_entry = |dir: &mut Vec<u8>,
                           name: &str,
                           obj: u8,
                           left: u32,
                           right: u32,
                           child: u32,
                           start: u32,
                           size: u64| {
            let mut e = vec![0u8; 128];
            let utf16: Vec<u16> = name.encode_utf16().collect();
            for (k, u) in utf16.iter().enumerate().take(31) {
                e[k * 2..k * 2 + 2].copy_from_slice(&u.to_le_bytes());
            }
            let name_len = ((utf16.len().min(31) + 1) * 2) as u16;
            e[64..66].copy_from_slice(&name_len.to_le_bytes());
            e[66] = obj;
            e[67] = 1; // black
            e[68..72].copy_from_slice(&left.to_le_bytes());
            e[72..76].copy_from_slice(&right.to_le_bytes());
            e[76..80].copy_from_slice(&child.to_le_bytes());
            e[116..120].copy_from_slice(&start.to_le_bytes());
            e[120..128].copy_from_slice(&size.to_le_bytes());
            dir.extend_from_slice(&e);
        };

        // Root entry (type 5) owns the mini stream; its child is the first stream
        // entry, and the streams form a right-leaning list in sorted order.
        let first_child = if ordered.is_empty() { NOSTREAM } else { 1 };
        write_entry(
            &mut dir,
            "Root Entry",
            5,
            NOSTREAM,
            NOSTREAM,
            first_child,
            ministream_start,
            mini_stream.len() as u64,
        );
        for (i, (name, data)) in ordered.iter().enumerate() {
            let right = if i + 1 < ordered.len() {
                (i + 2) as u32
            } else {
                NOSTREAM
            };
            write_entry(
                &mut dir,
                name,
                2,
                NOSTREAM,
                right,
                NOSTREAM,
                placement[i].0,
                data.len() as u64,
            );
        }
        dir.resize(round_up(dir.len(), SECTOR), 0);

        // Serialize the header.
        let mut out = vec![0u8; SECTOR];
        out[..8].copy_from_slice(&super::CFB_MAGIC);
        out[24..26].copy_from_slice(&0x003Eu16.to_le_bytes()); // minor
        out[26..28].copy_from_slice(&0x0003u16.to_le_bytes()); // major (v3, 512)
        out[28..30].copy_from_slice(&0xFFFEu16.to_le_bytes()); // byte order
        out[30..32].copy_from_slice(&9u16.to_le_bytes()); // sector shift
        out[32..34].copy_from_slice(&6u16.to_le_bytes()); // mini sector shift
        out[44..48].copy_from_slice(&(num_fat as u32).to_le_bytes());
        out[48..52].copy_from_slice(&dir_start.to_le_bytes());
        out[56..60].copy_from_slice(&(CUTOFF as u32).to_le_bytes());
        out[60..64].copy_from_slice(
            &(if minifat_sectors > 0 {
                minifat_start
            } else {
                ENDOFCHAIN
            })
            .to_le_bytes(),
        );
        out[64..68].copy_from_slice(&(minifat_sectors as u32).to_le_bytes());
        out[68..72].copy_from_slice(&ENDOFCHAIN.to_le_bytes()); // first DIFAT
        out[72..76].copy_from_slice(&0u32.to_le_bytes()); // num DIFAT
        for i in 0..109 {
            let v = if i < num_fat { i as u32 } else { FREESECT };
            out[76 + i * 4..80 + i * 4].copy_from_slice(&v.to_le_bytes());
        }

        // Sector payloads, in sector-number order.
        let mut fat_bytes = Vec::with_capacity(fat.len() * 4);
        for v in &fat {
            fat_bytes.extend_from_slice(&v.to_le_bytes());
        }
        out.extend_from_slice(&fat_bytes); // FAT sectors
        out.extend_from_slice(&dir); // directory
        if minifat_sectors > 0 {
            let mut mf = Vec::new();
            for v in &mini_fat {
                mf.extend_from_slice(&v.to_le_bytes());
            }
            mf.resize(minifat_sectors * SECTOR, {
                // pad remaining mini-FAT slots with FREESECT
                0xFF
            });
            out.extend_from_slice(&mf);
        }
        if ministream_sectors > 0 {
            let mut ms = mini_stream.clone();
            ms.resize(ministream_sectors * SECTOR, 0);
            out.extend_from_slice(&ms);
        }
        for (i, (_, d)) in ordered.iter().enumerate() {
            if d.len() >= CUTOFF {
                let mut b = d.to_vec();
                b.resize(big_sectors[i] * SECTOR, 0);
                out.extend_from_slice(&b);
            }
        }
        Ok(out)
    }
}

/// The hash used throughout an agile-encryption file.
#[derive(Clone, Copy)]
enum HashAlgo {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgo {
    fn parse(name: &str) -> Result<Self> {
        let n = name.to_ascii_uppercase().replace('-', "");
        Ok(match n.as_str() {
            "SHA1" => HashAlgo::Sha1,
            "SHA256" => HashAlgo::Sha256,
            "SHA384" => HashAlgo::Sha384,
            "SHA512" => HashAlgo::Sha512,
            _ => {
                return Err(RustypyxlError::InvalidFormat(format!(
                    "unsupported hash algorithm {name:?}"
                )))
            }
        })
    }

    fn hash(&self, data: &[u8]) -> Vec<u8> {
        match self {
            HashAlgo::Sha1 => Sha1::digest(data).to_vec(),
            HashAlgo::Sha256 => Sha256::digest(data).to_vec(),
            HashAlgo::Sha384 => Sha384::digest(data).to_vec(),
            HashAlgo::Sha512 => Sha512::digest(data).to_vec(),
        }
    }
}

/// Parsed `<keyData>` / `<encryptedKey>` attributes from the agile XML.
#[derive(Default)]
struct AgileParams {
    key_data_salt: Vec<u8>,
    key_data_hash: Option<HashAlgo>,
    block_size: usize,
    spin_count: u32,
    key_bits: usize,
    key_salt: Vec<u8>,
    key_hash: Option<HashAlgo>,
    encrypted_verifier_input: Vec<u8>,
    encrypted_verifier_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

fn decrypt_agile(xml: &[u8], package: &[u8], password: &str) -> Result<Vec<u8>> {
    let p = parse_agile_xml(xml)?;
    let key_hash = p
        .key_hash
        .ok_or_else(|| RustypyxlError::InvalidFormat("missing key hash algorithm".into()))?;
    let data_hash = p
        .key_data_hash
        .ok_or_else(|| RustypyxlError::InvalidFormat("missing keyData hash algorithm".into()))?;
    let key_bytes = p.key_bits / 8;
    if p.block_size == 0 || p.block_size > 16 {
        return Err(RustypyxlError::InvalidFormat(
            "invalid agile block size".into(),
        ));
    }

    let pw16: Vec<u8> = password
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect();

    // Verify the password before doing more work.
    let verifier_input_key = derive_key(
        key_hash,
        &pw16,
        &p.key_salt,
        p.spin_count,
        &BLOCK_KEY_VERIFIER_INPUT,
        key_bytes,
    );
    let verifier_input = aes_cbc_decrypt(
        &verifier_input_key,
        &p.key_salt[..p.block_size],
        &p.encrypted_verifier_input,
    )?;
    let verifier_value_key = derive_key(
        key_hash,
        &pw16,
        &p.key_salt,
        p.spin_count,
        &BLOCK_KEY_VERIFIER_VALUE,
        key_bytes,
    );
    let verifier_value = aes_cbc_decrypt(
        &verifier_value_key,
        &p.key_salt[..p.block_size],
        &p.encrypted_verifier_value,
    )?;
    let computed = key_hash.hash(&verifier_input);
    let hash_len = computed.len().min(verifier_value.len());
    if computed[..hash_len] != verifier_value[..hash_len] {
        return Err(RustypyxlError::InvalidFormat(
            "incorrect password for encrypted workbook".into(),
        ));
    }

    // Recover the package key.
    let key_value_key = derive_key(
        key_hash,
        &pw16,
        &p.key_salt,
        p.spin_count,
        &BLOCK_KEY_ENCRYPTED_KEY,
        key_bytes,
    );
    let mut secret = aes_cbc_decrypt(
        &key_value_key,
        &p.key_salt[..p.block_size],
        &p.encrypted_key_value,
    )?;
    secret.truncate(key_bytes);

    // Decrypt the package in 4096-byte segments; each segment's IV is
    // hash(keyDataSalt || segmentIndex) truncated to the block size.
    if package.len() < 8 {
        return Err(RustypyxlError::InvalidFormat(
            "truncated EncryptedPackage".into(),
        ));
    }
    let total_size = u64::from_le_bytes(package[..8].try_into().unwrap()) as usize;
    if total_size > package.len() {
        return Err(RustypyxlError::InvalidFormat(
            "encrypted package size prefix exceeds the stream".into(),
        ));
    }
    let body = &package[8..];
    let mut out = Vec::with_capacity(total_size);
    for (i, segment) in body.chunks(4096).enumerate() {
        let mut iv_input = p.key_data_salt.clone();
        iv_input.extend_from_slice(&(i as u32).to_le_bytes());
        let iv = data_hash.hash(&iv_input);
        // Pad the final segment up to a block boundary if needed.
        let mut seg = segment.to_vec();
        if !seg.len().is_multiple_of(16) {
            seg.resize(seg.len() + (16 - seg.len() % 16), 0);
        }
        let decrypted = aes_cbc_decrypt(&secret, &iv[..p.block_size], &seg)?;
        out.extend_from_slice(&decrypted);
    }
    out.truncate(total_size);
    Ok(out)
}

/// Derive a key from the password: an iterated hash (spinCount) mixed with a
/// per-purpose block key, truncated/padded to the required length.
fn derive_key(
    algo: HashAlgo,
    pw16: &[u8],
    salt: &[u8],
    spin_count: u32,
    block_key: &[u8],
    key_bytes: usize,
) -> Vec<u8> {
    let mut h = {
        let mut input = salt.to_vec();
        input.extend_from_slice(pw16);
        algo.hash(&input)
    };
    for i in 0..spin_count {
        let mut input = i.to_le_bytes().to_vec();
        input.extend_from_slice(&h);
        h = algo.hash(&input);
    }
    let mut input = h;
    input.extend_from_slice(block_key);
    let mut key = algo.hash(&input);
    key.resize(key_bytes, 0x36);
    key
}

/// AES-CBC decrypt (no padding removal). `data` must be a multiple of 16 bytes.
fn aes_cbc_decrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::generic_array::GenericArray;
    use aes::cipher::{BlockDecrypt, KeyInit};

    if data.is_empty() || !data.len().is_multiple_of(16) {
        return Err(RustypyxlError::InvalidFormat(
            "encrypted block is not a multiple of 16 bytes".into(),
        ));
    }
    if iv.len() < 16 {
        return Err(RustypyxlError::InvalidFormat(
            "IV shorter than a block".into(),
        ));
    }

    macro_rules! run {
        ($cipher:ty) => {{
            let cipher = <$cipher>::new(GenericArray::from_slice(key));
            let mut prev = [0u8; 16];
            prev.copy_from_slice(&iv[..16]);
            let mut out = Vec::with_capacity(data.len());
            for chunk in data.chunks(16) {
                let mut ciphertext = [0u8; 16];
                ciphertext.copy_from_slice(chunk);
                let mut block = GenericArray::clone_from_slice(chunk);
                cipher.decrypt_block(&mut block);
                for j in 0..16 {
                    block[j] ^= prev[j];
                }
                out.extend_from_slice(&block);
                prev = ciphertext;
            }
            out
        }};
    }

    let out = match key.len() {
        16 => run!(aes::Aes128),
        24 => run!(aes::Aes192),
        32 => run!(aes::Aes256),
        n => {
            return Err(RustypyxlError::InvalidFormat(format!(
                "unsupported AES key size {n}"
            )))
        }
    };
    Ok(out)
}

// ---------- encryption (writing) ----------

/// Block keys for the agile data-integrity HMAC.
#[cfg(feature = "encrypt")]
const BLOCK_KEY_HMAC_KEY: [u8; 8] = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
#[cfg(feature = "encrypt")]
const BLOCK_KEY_HMAC_VALUE: [u8; 8] = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

/// Encrypt a plaintext ZIP (a saved workbook) into an agile-encrypted CFB
/// container. Uses AES-256-CBC and SHA-512, matching what Office writes.
#[cfg(feature = "encrypt")]
pub fn encrypt(plain: &[u8], password: &str) -> Result<Vec<u8>> {
    let algo = HashAlgo::Sha512;
    let key_bytes = 32usize; // AES-256
    let block = 16usize;
    let spin = 100_000u32;

    let key_data_salt = random_bytes(16)?;
    let key_salt = random_bytes(16)?;
    let package_key = random_bytes(key_bytes)?;
    let pw16: Vec<u8> = password
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect();

    // Password verifier.
    let verifier_input = random_bytes(16)?;
    let vin_key = derive_key(
        algo,
        &pw16,
        &key_salt,
        spin,
        &BLOCK_KEY_VERIFIER_INPUT,
        key_bytes,
    );
    let enc_verifier_input = aes_cbc_encrypt(&vin_key, &key_salt, &verifier_input)?;
    let verifier_hash = algo.hash(&verifier_input);
    let vval_key = derive_key(
        algo,
        &pw16,
        &key_salt,
        spin,
        &BLOCK_KEY_VERIFIER_VALUE,
        key_bytes,
    );
    let enc_verifier_value = aes_cbc_encrypt(&vval_key, &key_salt, &pad_to_16(&verifier_hash))?;

    // Encrypted package key.
    let kv_key = derive_key(
        algo,
        &pw16,
        &key_salt,
        spin,
        &BLOCK_KEY_ENCRYPTED_KEY,
        key_bytes,
    );
    let enc_key_value = aes_cbc_encrypt(&kv_key, &key_salt, &package_key)?;

    // Encrypt the package in 4096-byte segments, prefixed by the plaintext size.
    let mut package = (plain.len() as u64).to_le_bytes().to_vec();
    let padded = pad_to_16(plain);
    for (i, segment) in padded.chunks(4096).enumerate() {
        let mut iv_input = key_data_salt.clone();
        iv_input.extend_from_slice(&(i as u32).to_le_bytes());
        let iv = algo.hash(&iv_input);
        package.extend_from_slice(&aes_cbc_encrypt(&package_key, &iv[..block], segment)?);
    }

    // Data integrity: an HMAC over the encrypted package, both key and value
    // themselves encrypted with the package key.
    let hmac_key = random_bytes(64)?;
    let iv_hk = {
        let mut v = key_data_salt.clone();
        v.extend_from_slice(&BLOCK_KEY_HMAC_KEY);
        algo.hash(&v)
    };
    let enc_hmac_key = aes_cbc_encrypt(&package_key, &iv_hk[..block], &pad_to_16(&hmac_key))?;
    let hmac_value = hmac_sha512(&hmac_key, &package);
    let iv_hv = {
        let mut v = key_data_salt.clone();
        v.extend_from_slice(&BLOCK_KEY_HMAC_VALUE);
        algo.hash(&v)
    };
    let enc_hmac_value = aes_cbc_encrypt(&package_key, &iv_hv[..block], &pad_to_16(&hmac_value))?;

    // Serialize the agile EncryptionInfo (an 8-byte version header + XML).
    let xml = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" "#,
            r#"xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" "#,
            r#"xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">"#,
            r#"<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" "#,
            r#"cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{kds}"/>"#,
            r#"<dataIntegrity encryptedHmacKey="{ehk}" encryptedHmacValue="{ehv}"/>"#,
            r#"<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#,
            r#"<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" "#,
            r#"cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{ks}" "#,
            r#"encryptedVerifierHashInput="{evi}" encryptedVerifierHashValue="{evv}" encryptedKeyValue="{ekv}"/>"#,
            r#"</keyEncryptor></keyEncryptors></encryption>"#
        ),
        kds = base64_encode(&key_data_salt),
        ehk = base64_encode(&enc_hmac_key),
        ehv = base64_encode(&enc_hmac_value),
        ks = base64_encode(&key_salt),
        evi = base64_encode(&enc_verifier_input),
        evv = base64_encode(&enc_verifier_value),
        ekv = base64_encode(&enc_key_value),
    );
    let mut info = vec![0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00];
    info.extend_from_slice(xml.as_bytes());

    cfb::write_container(&[("EncryptionInfo", &info), ("EncryptedPackage", &package)])
}

/// Fill a fresh buffer with cryptographically secure random bytes.
#[cfg(feature = "encrypt")]
fn random_bytes(n: usize) -> Result<Vec<u8>> {
    let mut buf = vec![0u8; n];
    getrandom::fill(&mut buf)
        .map_err(|e| RustypyxlError::Custom(format!("failed to gather randomness: {e}")))?;
    Ok(buf)
}

/// Zero-pad a byte slice up to the next 16-byte (AES block) boundary.
#[cfg(feature = "encrypt")]
fn pad_to_16(data: &[u8]) -> Vec<u8> {
    let mut v = data.to_vec();
    if !v.len().is_multiple_of(16) {
        v.resize(v.len() + (16 - v.len() % 16), 0);
    }
    v
}

/// HMAC-SHA512.
#[cfg(feature = "encrypt")]
fn hmac_sha512(key: &[u8], data: &[u8]) -> Vec<u8> {
    use hmac::{Mac, SimpleHmac};
    let mut mac = SimpleHmac::<Sha512>::new_from_slice(key).expect("HMAC accepts any key length");
    mac.update(data);
    mac.finalize().into_bytes().to_vec()
}

/// AES-CBC encrypt (no padding). `data` must be a multiple of 16 bytes.
#[cfg(feature = "encrypt")]
fn aes_cbc_encrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::generic_array::GenericArray;
    use aes::cipher::{BlockEncrypt, KeyInit};

    if !data.len().is_multiple_of(16) {
        return Err(RustypyxlError::InvalidFormat(
            "plaintext block is not a multiple of 16 bytes".into(),
        ));
    }
    macro_rules! run {
        ($cipher:ty) => {{
            let cipher = <$cipher>::new(GenericArray::from_slice(key));
            let mut prev = [0u8; 16];
            prev.copy_from_slice(&iv[..16]);
            let mut out = Vec::with_capacity(data.len());
            for chunk in data.chunks(16) {
                let mut block = [0u8; 16];
                for j in 0..16 {
                    block[j] = chunk[j] ^ prev[j];
                }
                let mut ga = GenericArray::clone_from_slice(&block);
                cipher.encrypt_block(&mut ga);
                prev.copy_from_slice(&ga);
                out.extend_from_slice(&ga);
            }
            out
        }};
    }
    let out = match key.len() {
        16 => run!(aes::Aes128),
        24 => run!(aes::Aes192),
        32 => run!(aes::Aes256),
        n => {
            return Err(RustypyxlError::InvalidFormat(format!(
                "unsupported AES key size {n}"
            )))
        }
    };
    Ok(out)
}

/// Standard-alphabet base64 encoder (mirrors the in-crate decoder).
#[cfg(feature = "encrypt")]
fn base64_encode(data: &[u8]) -> String {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(data.len().div_ceil(3) * 4);
    for chunk in data.chunks(3) {
        let b = [
            chunk[0],
            *chunk.get(1).unwrap_or(&0),
            *chunk.get(2).unwrap_or(&0),
        ];
        out.push(ALPHABET[(b[0] >> 2) as usize] as char);
        out.push(ALPHABET[(((b[0] & 0x03) << 4) | (b[1] >> 4)) as usize] as char);
        if chunk.len() > 1 {
            out.push(ALPHABET[(((b[1] & 0x0f) << 2) | (b[2] >> 6)) as usize] as char);
        } else {
            out.push('=');
        }
        if chunk.len() > 2 {
            out.push(ALPHABET[(b[2] & 0x3f) as usize] as char);
        } else {
            out.push('=');
        }
    }
    out
}

/// Parse the two agile-encryption elements out of the EncryptionInfo XML.
fn parse_agile_xml(xml: &[u8]) -> Result<AgileParams> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut p = AgileParams::default();

    let attr = |e: &quick_xml::events::BytesStart, key: &[u8]| -> Option<String> {
        e.attributes()
            .flatten()
            .find(|a| a.key.local_name().as_ref() == key)
            .and_then(|a| a.unescape_value().ok().map(|v| v.into_owned()))
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"keyData" => {
                    if let Some(s) = attr(&e, b"saltValue") {
                        p.key_data_salt = base64_decode(&s)?;
                    }
                    if let Some(h) = attr(&e, b"hashAlgorithm") {
                        p.key_data_hash = Some(HashAlgo::parse(&h)?);
                    }
                    if let Some(b) = attr(&e, b"blockSize").and_then(|v| v.parse().ok()) {
                        p.block_size = b;
                    }
                }
                b"encryptedKey" => {
                    if let Some(s) = attr(&e, b"spinCount").and_then(|v| v.parse().ok()) {
                        p.spin_count = s;
                    }
                    if let Some(k) = attr(&e, b"keyBits").and_then(|v| v.parse().ok()) {
                        p.key_bits = k;
                    }
                    if let Some(h) = attr(&e, b"hashAlgorithm") {
                        p.key_hash = Some(HashAlgo::parse(&h)?);
                    }
                    if let Some(s) = attr(&e, b"saltValue") {
                        p.key_salt = base64_decode(&s)?;
                    }
                    if let Some(s) = attr(&e, b"encryptedVerifierHashInput") {
                        p.encrypted_verifier_input = base64_decode(&s)?;
                    }
                    if let Some(s) = attr(&e, b"encryptedVerifierHashValue") {
                        p.encrypted_verifier_value = base64_decode(&s)?;
                    }
                    if let Some(s) = attr(&e, b"encryptedKeyValue") {
                        p.encrypted_key_value = base64_decode(&s)?;
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(RustypyxlError::InvalidFormat(format!(
                    "malformed EncryptionInfo XML: {e}"
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    if p.key_bits == 0 || p.key_salt.is_empty() || p.encrypted_key_value.is_empty() {
        return Err(RustypyxlError::InvalidFormat(
            "incomplete agile encryption parameters".into(),
        ));
    }
    Ok(p)
}

/// Minimal standard-alphabet base64 decoder (avoids a base64 dependency).
fn base64_decode(s: &str) -> Result<Vec<u8>> {
    fn val(c: u8) -> Option<u8> {
        match c {
            b'A'..=b'Z' => Some(c - b'A'),
            b'a'..=b'z' => Some(c - b'a' + 26),
            b'0'..=b'9' => Some(c - b'0' + 52),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None,
        }
    }
    let mut out = Vec::with_capacity(s.len() * 3 / 4);
    let mut buf: u32 = 0;
    let mut bits = 0;
    for &c in s.trim().as_bytes() {
        if matches!(c, b'=' | b'\n' | b'\r' | b' ' | b'\t') {
            continue;
        }
        let v = val(c).ok_or_else(|| {
            RustypyxlError::InvalidFormat("invalid base64 in EncryptionInfo".into())
        })?;
        buf = (buf << 6) | v as u32;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((buf >> bits) as u8);
        }
    }
    Ok(out)
}
