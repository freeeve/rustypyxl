use std::env;
use std::fs::File;
use std::io::Write;
use std::time::Instant;

use pprof::protos::Message;
use pprof::ProfilerGuard;
use rustypyxl_core::Workbook;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!(
            "Usage: profile_read <file.xlsx> [flamegraph.svg] [profile.pb]"
        );
        eprintln!("Example: cargo run -p rustypyxl-core --features pprof --bin profile_read -- test.xlsx");
        std::process::exit(2);
    }

    let path = &args[1];
    let flamegraph_path = args.get(2).map(String::as_str).unwrap_or("flamegraph.svg");
    let pprof_path = args.get(3).map(String::as_str).unwrap_or("profile.pb");

    eprintln!("Starting profiler...");
    let guard = match ProfilerGuard::new(100) {
        Ok(guard) => guard,
        Err(err) => {
            eprintln!("pprof not available: {err}");
            std::process::exit(1);
        }
    };
    let start = Instant::now();
    let wb = Workbook::load(path).expect("failed to load workbook");
    let _ = wb.worksheets();
    let elapsed = start.elapsed();

    eprintln!("Read completed in {:.3}s", elapsed.as_secs_f64());

    let report = guard.report().build().expect("failed to build profile");

    let mut flamegraph_file =
        File::create(flamegraph_path).expect("failed to create flamegraph file");
    report
        .flamegraph(&mut flamegraph_file)
        .expect("failed to write flamegraph");

    let profile = report.pprof().expect("failed to build pprof");
    let mut body = Vec::new();
    profile
        .write_to_vec(&mut body)
        .expect("failed to encode pprof profile");

    let mut pprof_file = File::create(pprof_path).expect("failed to create pprof file");
    pprof_file
        .write_all(&body)
        .expect("failed to write pprof file");

    eprintln!("Wrote {flamegraph_path} and {pprof_path}");
}
