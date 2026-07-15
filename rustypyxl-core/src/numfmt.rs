//! Render a value under an Excel number-format code to the string Excel would
//! display for it. Excel stores only the format *code* (e.g. `"0.00%"`,
//! `"yyyy-mm-dd"`, `"#,##0.00;[Red](#,##0.00)"`); this module interprets that
//! code so readers can show data the way Excel does without re-implementing the
//! rules per caller.
//!
//! Scope is the common subset: the positive;negative;zero;text sections, digit
//! placeholders (`0` `#` `?`), thousands grouping and decimals, `%`, scaling
//! commas, quoted/escaped literals, currency (`[$…]`), and date/time tokens
//! (`yyyy mm dd hh ss AM/PM` plus elapsed `[h] [m] [s]`). Exotic constructs
//! (fractions `?/?`, scientific `E+`, literals interleaved between digits like
//! SSN masks) are not covered and fall back to a best effort.

use crate::cell::CellValue;

/// Map a built-in number-format id (0-49) to its implied format code. Ids
/// outside the built-in range have no implied code (they reference a custom
/// format defined in the styles part).
pub fn builtin_format_code(id: u32) -> Option<&'static str> {
    let code = match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "$#,##0_);($#,##0)",
        6 => "$#,##0_);[Red]($#,##0)",
        7 => "$#,##0.00_);($#,##0.00)",
        8 => "$#,##0.00_);[Red]($#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    };
    Some(code)
}

/// Format a cell value under a format code, returning the display string Excel
/// would show. Strings pass through the text section (`@`); booleans render as
/// TRUE/FALSE; empty cells render as the empty string.
pub fn format_value(value: &CellValue, code: &str) -> String {
    match value {
        CellValue::Number(n) => format_number(*n, code),
        CellValue::Boolean(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        CellValue::String(s) => format_text(s, code),
        CellValue::Date(s) => s.clone(),
        CellValue::Formula(f) => f.clone(),
        CellValue::Empty => String::new(),
    }
}

/// Apply the text section (`@`) of a format to a string. If the format has no
/// text section, the string is returned unchanged.
fn format_text(text: &str, code: &str) -> String {
    let sections = split_sections(code);
    // The text section is the 4th; a single-section code that contains `@` is
    // itself the text section. Otherwise text is shown verbatim.
    let section = if let Some(s) = sections.get(3) {
        s.clone()
    } else if sections.len() == 1 && sections[0].contains('@') {
        sections[0].clone()
    } else {
        return text.to_string();
    };
    let section = &section;
    let mut out = String::new();
    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '@' => out.push_str(text),
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                    out.push(q);
                }
            }
            '\\' => {
                if let Some(n) = chars.next() {
                    out.push(n);
                }
            }
            '[' => {
                // color / condition bracket, skip
                for b in chars.by_ref() {
                    if b == ']' {
                        break;
                    }
                }
            }
            _ => out.push(c),
        }
    }
    out
}

/// Format a numeric value under a format code.
pub fn format_number(value: f64, code: &str) -> String {
    if code.is_empty() || code.eq_ignore_ascii_case("general") {
        return format_general(value);
    }

    let sections = split_sections(code);
    let (section, force_negative) = choose_section(&sections, value);

    if is_datetime_section(&section) {
        return render_datetime(value, &section);
    }

    let mut out = String::new();
    if force_negative && value != 0.0 {
        out.push('-');
    }
    out.push_str(&render_numeric(value.abs(), &section));
    out
}

/// Excel's General format: shortest round-tripping representation, integers
/// without a decimal point.
fn format_general(value: f64) -> String {
    if value == value.trunc() && value.abs() < 1e15 {
        return format!("{}", value as i64);
    }
    let s = format!("{}", value);
    s
}

/// Split a format code into its `;`-separated sections, respecting quotes and
/// bracketed groups so a `;` inside them does not split.
fn split_sections(code: &str) -> Vec<String> {
    let mut sections = Vec::new();
    let mut cur = String::new();
    let mut chars = code.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            ';' => {
                sections.push(std::mem::take(&mut cur));
            }
            '"' => {
                cur.push(c);
                for q in chars.by_ref() {
                    cur.push(q);
                    if q == '"' {
                        break;
                    }
                }
            }
            '[' => {
                cur.push(c);
                for b in chars.by_ref() {
                    cur.push(b);
                    if b == ']' {
                        break;
                    }
                }
            }
            '\\' => {
                cur.push(c);
                if let Some(n) = chars.next() {
                    cur.push(n);
                }
            }
            _ => cur.push(c),
        }
    }
    sections.push(cur);
    sections
}

/// Pick the section that applies to a value and whether a leading minus must be
/// synthesized (only when a negative value falls through to the positive
/// section because no dedicated negative section exists).
fn choose_section(sections: &[String], value: f64) -> (String, bool) {
    let non_text: Vec<&String> = sections.iter().take(3).collect();
    match non_text.len() {
        0 => (String::new(), value < 0.0),
        1 => (non_text[0].clone(), value < 0.0),
        2 => {
            if value < 0.0 {
                (non_text[1].clone(), false)
            } else {
                (non_text[0].clone(), false)
            }
        }
        _ => {
            if value < 0.0 {
                (non_text[1].clone(), false)
            } else if value == 0.0 {
                (non_text[2].clone(), false)
            } else {
                (non_text[0].clone(), false)
            }
        }
    }
}

/// A token from a numeric-format section.
enum NumTok {
    /// Digit placeholder: '0', '#', or '?'.
    Ph(char),
    /// Literal character to emit verbatim.
    Lit(char),
    /// Thousands separator / scaling comma.
    Comma,
    /// Decimal point.
    Dot,
    /// Percent (scales the value by 100).
    Percent,
}

/// Tokenize a numeric section, unwrapping quotes, escapes, underscores (which
/// reserve the width of the next char -> a space), currency brackets, and color
/// brackets.
fn tokenize_numeric(section: &str) -> Vec<NumTok> {
    let mut toks = Vec::new();
    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '0' | '#' | '?' => toks.push(NumTok::Ph(c)),
            ',' => toks.push(NumTok::Comma),
            '.' => toks.push(NumTok::Dot),
            '%' => toks.push(NumTok::Percent),
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                    toks.push(NumTok::Lit(q));
                }
            }
            '\\' => {
                if let Some(n) = chars.next() {
                    toks.push(NumTok::Lit(n));
                }
            }
            '_' => {
                // reserve width of the next char: render as a space
                chars.next();
                toks.push(NumTok::Lit(' '));
            }
            '*' => {
                // fill with the next char: not reproducible as a fixed string, skip it
                chars.next();
            }
            '[' => {
                let mut inner = String::new();
                for b in chars.by_ref() {
                    if b == ']' {
                        break;
                    }
                    inner.push(b);
                }
                // [$symbol-locale] contributes the currency symbol as a literal.
                if let Some(rest) = inner.strip_prefix('$') {
                    let symbol = rest.split('-').next().unwrap_or("");
                    for ch in symbol.chars() {
                        toks.push(NumTok::Lit(ch));
                    }
                }
                // color / condition brackets contribute nothing to the string.
            }
            other => toks.push(NumTok::Lit(other)),
        }
    }
    toks
}

/// Render a non-negative number through a numeric section's tokens.
fn render_numeric(value: f64, section: &str) -> String {
    let toks = tokenize_numeric(section);

    // Structural pass: percent scaling, decimal placeholder count, thousands
    // grouping, and trailing-comma scaling.
    let percent = toks.iter().filter(|t| matches!(t, NumTok::Percent)).count();
    let dot_idx = toks.iter().position(|t| matches!(t, NumTok::Dot));
    let int_end = dot_idx.unwrap_or(toks.len());

    let frac_places = toks
        .iter()
        .skip(dot_idx.map(|i| i + 1).unwrap_or(toks.len()))
        .filter(|t| matches!(t, NumTok::Ph(_)))
        .count();

    // Grouping if a comma sits between two integer-region digit placeholders.
    let int_ph_positions: Vec<usize> = toks
        .iter()
        .take(int_end)
        .enumerate()
        .filter(|(_, t)| matches!(t, NumTok::Ph(_)))
        .map(|(i, _)| i)
        .collect();
    let grouping = toks.iter().take(int_end).enumerate().any(|(i, t)| {
        matches!(t, NumTok::Comma)
            && int_ph_positions.first().is_some_and(|&f| i > f)
            && int_ph_positions.last().is_some_and(|&l| i < l)
    });
    // Trailing commas after the last integer placeholder scale the value down.
    let trailing_commas = match int_ph_positions.last() {
        Some(&last) => toks
            .iter()
            .take(int_end)
            .skip(last + 1)
            .filter(|t| matches!(t, NumTok::Comma))
            .count(),
        None => 0,
    };

    let mut scaled = value;
    for _ in 0..percent {
        scaled *= 100.0;
    }
    for _ in 0..trailing_commas {
        scaled /= 1000.0;
    }

    // Split into integer and fractional digit strings, rounded to frac_places.
    // Excel rounds half away from zero; pre-round with f64::round (which does
    // the same) so we don't inherit Rust's round-half-to-even in `{:.*}`.
    let factor = 10f64.powi(frac_places as i32);
    let scaled = (scaled * factor).round() / factor;
    let rounded = format!("{:.*}", frac_places, scaled);
    let (int_digits, frac_digits) = match rounded.split_once('.') {
        Some((i, f)) => (i.to_string(), f.to_string()),
        None => (rounded, String::new()),
    };

    let zeros_int = toks
        .iter()
        .take(int_end)
        .filter(|t| matches!(t, NumTok::Ph('0')))
        .count();

    let int_display = build_int_display(&int_digits, zeros_int, grouping);

    // Fractional placeholders, right to left: '#' trims trailing zeros, '?'
    // pads with a space, '0' always shows.
    let frac_phs: Vec<char> = toks
        .iter()
        .skip(dot_idx.map(|i| i + 1).unwrap_or(toks.len()))
        .filter_map(|t| match t {
            NumTok::Ph(c) => Some(*c),
            _ => None,
        })
        .collect();
    let frac_display = build_frac_display(&frac_digits, &frac_phs);

    assemble(&toks, dot_idx, &int_display, &frac_display)
}

/// Left-pad to the required minimum integer digits, drop a lone leading zero
/// when the format has no `0` there, then group in threes.
fn build_int_display(int_digits: &str, zeros_int: usize, grouping: bool) -> String {
    let mut digits = int_digits.to_string();
    if digits == "0" && zeros_int == 0 {
        // No mandatory integer digit and the value's integer part is zero: Excel
        // shows nothing (e.g. ".5" under "#.0" would show for 0.5, but we keep a
        // leading dot handled by the caller).
        digits.clear();
    }
    while digits.len() < zeros_int {
        digits.insert(0, '0');
    }
    if grouping && !digits.is_empty() {
        digits = group_thousands(&digits);
    }
    digits
}

/// Insert thousands separators into a run of digits.
fn group_thousands(digits: &str) -> String {
    let bytes = digits.as_bytes();
    let mut out = String::with_capacity(digits.len() + digits.len() / 3);
    let len = bytes.len();
    for (i, b) in bytes.iter().enumerate() {
        if i > 0 && (len - i).is_multiple_of(3) {
            out.push(',');
        }
        out.push(*b as char);
    }
    out
}

/// Build the fractional digits string honoring per-placeholder rules.
fn build_frac_display(frac_digits: &str, frac_phs: &[char]) -> String {
    if frac_phs.is_empty() {
        return String::new();
    }
    let digits: Vec<char> = frac_digits.chars().collect();
    // Determine the last index to keep: trailing insignificant zeros under a
    // '#' or '?' placeholder are dropped ('#' -> nothing, '?' -> a space); a '0'
    // placeholder always keeps its digit and stops the trim.
    let mut last: isize = frac_phs.len() as isize - 1;
    while last >= 0 {
        let i = last as usize;
        let ph = frac_phs[i];
        let d = digits.get(i).copied().unwrap_or('0');
        if (ph == '#' || ph == '?') && d == '0' {
            last -= 1;
        } else {
            break;
        }
    }
    let mut out = String::new();
    for (i, &ph) in frac_phs.iter().enumerate() {
        if (i as isize) <= last {
            out.push(digits.get(i).copied().unwrap_or('0'));
        } else if ph == '?' {
            out.push(' ');
        }
        // '#' beyond last: nothing. '0' beyond last cannot occur (trim stops on it).
    }
    out
}

/// Walk the tokens, substituting the prepared integer and fractional strings
/// for their respective placeholder runs and emitting literals in place.
fn assemble(
    toks: &[NumTok],
    dot_idx: Option<usize>,
    int_display: &str,
    frac_display: &str,
) -> String {
    let mut out = String::new();
    let mut int_field_done = false;
    let mut frac_field_done = false;
    for (i, tok) in toks.iter().enumerate() {
        let in_frac = dot_idx.is_some_and(|d| i > d);
        match tok {
            NumTok::Ph(_) | NumTok::Comma => {
                if in_frac {
                    if !frac_field_done {
                        out.push_str(frac_display);
                        frac_field_done = true;
                    }
                } else if !int_field_done {
                    out.push_str(int_display);
                    int_field_done = true;
                }
            }
            NumTok::Dot => {
                // Only emit the decimal point if fractional digits will follow.
                if !frac_display.is_empty() {
                    out.push('.');
                }
            }
            NumTok::Percent => out.push('%'),
            NumTok::Lit(c) => out.push(*c),
        }
    }
    out
}

// ------- date / time -------

/// Whether a section should be interpreted as a date/time rather than a number.
fn is_datetime_section(section: &str) -> bool {
    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' => {
                chars.next();
            }
            '[' => {
                // elapsed-time brackets [h] [m] [s] make this a time format
                let mut inner = String::new();
                for b in chars.by_ref() {
                    if b == ']' {
                        break;
                    }
                    inner.push(b);
                }
                let low = inner.to_ascii_lowercase();
                if low.chars().all(|c| c == 'h' || c == 'm' || c == 's') && !low.is_empty() {
                    return true;
                }
            }
            'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => return true,
            'm' | 'M' => return true,
            'a' | 'A' => {
                // AM/PM marker
                return true;
            }
            _ => {}
        }
    }
    false
}

/// A broken-down calendar date/time derived from an Excel serial.
struct DateParts {
    year: i64,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    weekday: u32, // 0 = Sunday
}

/// Days from civil date (proleptic Gregorian), after Howard Hinnant.
fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let doy = (153 * (if m > 2 { m - 3 } else { m + 9 }) as i64 + 2) / 5 + d as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

/// Civil date from a day count (proleptic Gregorian), after Howard Hinnant.
fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = (if mp < 10 { mp + 3 } else { mp - 9 }) as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

/// Convert an Excel serial date to calendar parts, honoring the 1900 date
/// system and its fictitious 1900-02-29 leap day.
fn serial_to_parts(serial: f64) -> DateParts {
    let whole = serial.floor() as i64;
    // Excel serial 1 = 1900-01-01. Serials >= 60 are shifted by one because
    // Excel counts a nonexistent 1900-02-29.
    let adjusted = if whole >= 60 { whole - 1 } else { whole };
    let base = days_from_civil(1899, 12, 31);
    let abs_days = base + adjusted; // days since 1970-01-01 (civil-function epoch)
    let (year, month, day) = civil_from_days(abs_days);
    // 1970-01-01 was a Thursday, which is index 4 with Sunday = 0.
    let weekday = (abs_days + 4).rem_euclid(7) as u32;

    let frac = serial - serial.floor();
    let total_seconds = (frac * 86400.0).round() as i64;
    let hour = (total_seconds / 3600).rem_euclid(24) as u32;
    let minute = ((total_seconds % 3600) / 60) as u32;
    let second = (total_seconds % 60) as u32;

    DateParts {
        year,
        month,
        day,
        hour,
        minute,
        second,
        weekday,
    }
}

const MONTHS_SHORT: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];
const MONTHS_LONG: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];
const WEEKDAYS_SHORT: [&str; 7] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const WEEKDAYS_LONG: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

/// Render a serial value through a date/time section.
fn render_datetime(serial: f64, section: &str) -> String {
    let p = serial_to_parts(serial);
    let has_ampm = section_has_ampm(section);

    let mut out = String::new();
    let chars: Vec<char> = section.chars().collect();
    let mut i = 0;

    // Track whether the previous emitted token was hours, so a following m/mm
    // renders as minutes rather than months.
    let mut prev_was_hour = false;

    while i < chars.len() {
        let c = chars[i];
        let lc = c.to_ascii_lowercase();
        match lc {
            '"' => {
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    out.push(chars[i]);
                    i += 1;
                }
                i += 1;
            }
            '\\' => {
                i += 1;
                if i < chars.len() {
                    out.push(chars[i]);
                    i += 1;
                }
            }
            '[' => {
                // elapsed-time token
                let start = i + 1;
                let mut j = start;
                while j < chars.len() && chars[j] != ']' {
                    j += 1;
                }
                let inner: String = chars[start..j].iter().collect::<String>().to_lowercase();
                let total_seconds = (serial * 86400.0).round() as i64;
                if inner.starts_with('h') {
                    out.push_str(&format!("{}", total_seconds / 3600));
                } else if inner.starts_with('m') {
                    out.push_str(&format!("{}", total_seconds / 60));
                } else if inner.starts_with('s') {
                    out.push_str(&format!("{}", total_seconds));
                }
                prev_was_hour = inner.starts_with('h');
                i = j + 1;
            }
            'y' => {
                let n = run_len(&chars, i, 'y');
                if n >= 4 {
                    out.push_str(&format!("{:04}", p.year));
                } else {
                    out.push_str(&format!("{:02}", p.year.rem_euclid(100)));
                }
                i += n;
                prev_was_hour = false;
            }
            'm' => {
                let n = run_len(&chars, i, 'm');
                let next_is_seconds = next_token_is(&chars, i + n, 's');
                if prev_was_hour || next_is_seconds {
                    // minutes
                    match n {
                        1 => out.push_str(&format!("{}", p.minute)),
                        _ => out.push_str(&format!("{:02}", p.minute)),
                    }
                } else {
                    match n {
                        1 => out.push_str(&format!("{}", p.month)),
                        2 => out.push_str(&format!("{:02}", p.month)),
                        3 => out.push_str(MONTHS_SHORT[(p.month - 1) as usize]),
                        _ => out.push_str(MONTHS_LONG[(p.month - 1) as usize]),
                    }
                }
                i += n;
                prev_was_hour = false;
            }
            'd' => {
                let n = run_len(&chars, i, 'd');
                match n {
                    1 => out.push_str(&format!("{}", p.day)),
                    2 => out.push_str(&format!("{:02}", p.day)),
                    3 => out.push_str(WEEKDAYS_SHORT[p.weekday as usize]),
                    _ => out.push_str(WEEKDAYS_LONG[p.weekday as usize]),
                }
                i += n;
                prev_was_hour = false;
            }
            'h' => {
                let n = run_len(&chars, i, 'h');
                let hour = if has_ampm {
                    let h = p.hour % 12;
                    if h == 0 {
                        12
                    } else {
                        h
                    }
                } else {
                    p.hour
                };
                if n >= 2 {
                    out.push_str(&format!("{:02}", hour));
                } else {
                    out.push_str(&format!("{}", hour));
                }
                i += n;
                prev_was_hour = true;
            }
            's' => {
                let n = run_len(&chars, i, 's');
                if n >= 2 {
                    out.push_str(&format!("{:02}", p.second));
                } else {
                    out.push_str(&format!("{}", p.second));
                }
                i += n;
                prev_was_hour = false;
            }
            'a' => {
                // AM/PM or A/P
                if matches_ci(&chars, i, "am/pm") {
                    out.push_str(if p.hour < 12 { "AM" } else { "PM" });
                    i += 5;
                } else if matches_ci(&chars, i, "a/p") {
                    out.push_str(if p.hour < 12 { "A" } else { "P" });
                    i += 3;
                } else {
                    out.push(c);
                    i += 1;
                }
                prev_was_hour = false;
            }
            _ => {
                out.push(c);
                i += 1;
            }
        }
    }
    out
}

/// Number of consecutive `target` chars (case-insensitive) starting at `i`.
fn run_len(chars: &[char], i: usize, target: char) -> usize {
    let mut n = 0;
    while i + n < chars.len() && chars[i + n].to_ascii_lowercase() == target {
        n += 1;
    }
    n
}

/// Whether the next format token (skipping `:` and spaces) begins with `target`.
fn next_token_is(chars: &[char], mut i: usize, target: char) -> bool {
    while i < chars.len() && (chars[i] == ':' || chars[i] == ' ') {
        i += 1;
    }
    i < chars.len() && chars[i].to_ascii_lowercase() == target
}

/// Whether the section contains an AM/PM (or A/P) marker.
fn section_has_ampm(section: &str) -> bool {
    let low = section.to_ascii_lowercase();
    low.contains("am/pm") || low.contains("a/p")
}

/// Case-insensitive match of `needle` at position `i` in `chars`.
fn matches_ci(chars: &[char], i: usize, needle: &str) -> bool {
    let nchars: Vec<char> = needle.chars().collect();
    if i + nchars.len() > chars.len() {
        return false;
    }
    for (k, nc) in nchars.iter().enumerate() {
        if chars[i + k].to_ascii_lowercase() != *nc {
            return false;
        }
    }
    true
}

#[cfg(test)]
mod tests {
    use super::*;

    fn f(value: f64, code: &str) -> String {
        format_number(value, code)
    }

    #[test]
    fn plain_integer_and_decimals() {
        assert_eq!(f(1234.0, "0"), "1234");
        assert_eq!(f(1234.5, "0"), "1235"); // rounds
        assert_eq!(f(1234.5, "0.00"), "1234.50");
        assert_eq!(f(0.5, "0.00"), "0.50");
        assert_eq!(f(0.5, "#.##"), ".5");
        assert_eq!(f(5.0, "0.##"), "5");
    }

    #[test]
    fn thousands_grouping() {
        assert_eq!(f(1234567.0, "#,##0"), "1,234,567");
        assert_eq!(f(1234567.891, "#,##0.00"), "1,234,567.89");
        assert_eq!(f(12.0, "#,##0"), "12");
    }

    #[test]
    fn percent() {
        assert_eq!(f(0.1234, "0%"), "12%");
        assert_eq!(f(0.1234, "0.00%"), "12.34%");
    }

    #[test]
    fn negative_sections() {
        assert_eq!(f(-5.0, "0"), "-5");
        assert_eq!(f(-1234.0, "#,##0;(#,##0)"), "(1,234)");
        assert_eq!(f(1234.0, "#,##0;(#,##0)"), "1,234");
        assert_eq!(f(0.0, "0.00;(0.00);\"zero\""), "zero");
    }

    #[test]
    fn currency_literal() {
        assert_eq!(f(1234.5, "$#,##0.00"), "$1,234.50");
        assert_eq!(f(1234.5, "[$$-409]#,##0.00"), "$1,234.50");
    }

    #[test]
    fn scaling_commas() {
        assert_eq!(f(1_500_000.0, "#,##0,"), "1,500");
        assert_eq!(f(1_500_000_000.0, "#,##0,,"), "1,500");
    }

    #[test]
    fn dates() {
        // 2023-01-15 is Excel serial 44941.
        assert_eq!(f(44941.0, "yyyy-mm-dd"), "2023-01-15");
        assert_eq!(f(44941.0, "m/d/yy"), "1/15/23");
        assert_eq!(f(44941.0, "d-mmm-yyyy"), "15-Jan-2023");
        assert_eq!(f(44941.0, "mmmm"), "January");
        assert_eq!(f(44941.0, "dddd"), "Sunday");
    }

    #[test]
    fn times() {
        // 0.5 = noon, 0.25 = 6am
        assert_eq!(f(0.5, "h:mm"), "12:00");
        assert_eq!(f(0.25, "h:mm AM/PM"), "6:00 AM");
        assert_eq!(f(0.75, "h:mm AM/PM"), "6:00 PM");
        assert_eq!(f(44941.5, "yyyy-mm-dd hh:mm"), "2023-01-15 12:00");
    }

    #[test]
    fn minutes_vs_months() {
        // m after h is minutes; standalone is month
        assert_eq!(f(44941.5, "hh:mm"), "12:00");
        assert_eq!(f(44941.0, "mm"), "01");
    }

    #[test]
    fn elapsed_time() {
        // 1.5 days = 36 hours
        assert_eq!(f(1.5, "[h]:mm"), "36:00");
    }

    #[test]
    fn builtins() {
        assert_eq!(builtin_format_code(0), Some("General"));
        assert_eq!(builtin_format_code(2), Some("0.00"));
        assert_eq!(builtin_format_code(9), Some("0%"));
        assert_eq!(builtin_format_code(999), None);
    }

    #[test]
    fn general_and_text() {
        assert_eq!(format_number(42.0, "General"), "42");
        assert_eq!(format_number(42.5, "General"), "42.5");
        assert_eq!(format_value(&CellValue::from("hi"), "@"), "hi");
        assert_eq!(format_value(&CellValue::from("hi"), "\"<\"@\">\""), "<hi>");
    }
}

#[cfg(test)]
mod coverage_tests {
    use super::*;

    #[test]
    fn all_builtin_codes_resolve_or_none() {
        // Every defined id renders something; gaps return None.
        for id in [
            1u32, 3, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 45, 46, 48,
        ] {
            assert!(builtin_format_code(id).is_some(), "id {id}");
        }
        assert_eq!(builtin_format_code(50), None);
    }

    #[test]
    fn value_variants() {
        assert_eq!(format_value(&CellValue::Boolean(true), "0"), "TRUE");
        assert_eq!(format_value(&CellValue::Boolean(false), "0"), "FALSE");
        assert_eq!(format_value(&CellValue::Empty, "0"), "");
        assert_eq!(
            format_value(&CellValue::Date("2023-01-01".into()), "yyyy"),
            "2023-01-01"
        );
        assert_eq!(
            format_value(&CellValue::Formula("SUM(A1)".into()), "0"),
            "SUM(A1)"
        );
        assert_eq!(format_value(&CellValue::Number(5.0), "0"), "5");
    }

    #[test]
    fn general_negatives_and_empty_code() {
        assert_eq!(format_number(-42.0, "General"), "-42");
        assert_eq!(format_number(3.5, ""), "3.5");
        assert_eq!(format_number(1e20, "General"), format!("{}", 1e20f64));
    }

    #[test]
    fn zero_section_and_color_brackets() {
        // color bracket in the negative section is stripped from the output
        assert_eq!(format_number(-5.0, "0.0;[Red]-0.0"), "-5.0");
        // explicit zero section
        assert_eq!(format_number(0.0, "0;;\"--\""), "--");
    }

    #[test]
    fn question_mark_padding_and_month_names() {
        // '?' shows a space for insignificant trailing positions (decimal align)
        assert_eq!(format_number(1.5, "0.0??"), "1.5  ");
        // a significant digit under '?' still shows
        assert_eq!(format_number(1.25, "0.0??"), "1.25 ");
        assert_eq!(format_number(44941.0, "mmm"), "Jan");
        assert_eq!(format_number(44941.0, "ddd"), "Sun");
        assert_eq!(format_number(44941.0, "yy"), "23");
    }

    #[test]
    fn elapsed_minutes_and_seconds() {
        // 1.5 days = 2160 minutes, 129600 seconds
        assert_eq!(format_number(1.5, "[m]"), "2160");
        assert_eq!(format_number(1.5, "[s]"), "129600");
    }

    #[test]
    fn currency_and_literal_quotes() {
        assert_eq!(format_number(9.0, "\"$\"0.00"), "$9.00");
        assert_eq!(format_number(5.0, "0\" units\""), "5 units");
    }
}
