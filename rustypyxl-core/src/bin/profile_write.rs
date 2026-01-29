use std::env;
use std::fs::File;
use std::io::Write;
use std::time::Instant;

use pprof::protos::Message;
use pprof::ProfilerGuard;
use rustypyxl_core::{Workbook, CellValue};

fn main() {
    let args: Vec<String> = env::args().collect();

    let rows: u32 = args.get(1).and_then(|s| s.parse().ok()).unwrap_or(10000);
    let cols: u32 = args.get(2).and_then(|s| s.parse().ok()).unwrap_or(20);
    let output_path = args.get(3).map(String::as_str).unwrap_or("profile_output.xlsx");
    let flamegraph_path = args.get(4).map(String::as_str).unwrap_or("flamegraph_write.svg");
    let pprof_path = args.get(5).map(String::as_str).unwrap_or("profile_write.pb");

    eprintln!("Creating workbook with {}x{} = {} cells...", rows, cols, rows * cols);

    // Create workbook with data (outside profiler)
    let mut wb = Workbook::new();
    wb.create_sheet(Some("Sheet1".to_string())).unwrap();

    for row in 1..=rows {
        for col in 1..=cols {
            let value = CellValue::String(format!("R{}C{}", row, col).into());
            wb.set_cell_value_in_sheet("Sheet1", row, col, value).unwrap();
        }
    }
    eprintln!("Data created, starting profiler for save...");

    let guard = match ProfilerGuard::new(100) {
        Ok(guard) => guard,
        Err(err) => {
            eprintln!("pprof not available: {err}");
            std::process::exit(1);
        }
    };

    let start = Instant::now();
    wb.save(output_path).expect("failed to save workbook");
    let elapsed = start.elapsed();

    eprintln!("Save completed in {:.3}s", elapsed.as_secs_f64());
    eprintln!("Throughput: {:.0} cells/sec", (rows * cols) as f64 / elapsed.as_secs_f64());

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

    // Clean up
    std::fs::remove_file(output_path).ok();
}
