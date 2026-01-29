use criterion::{black_box, criterion_group, criterion_main, Criterion, BenchmarkId};
use rustypyxl_core::workbook::Workbook;
use rustypyxl_core::worksheet::Worksheet;
use rustypyxl_core::cell::CellValue;
use std::collections::HashMap;

fn create_large_workbook(rows: u32, cols: u32) -> Workbook {
    let mut workbook = Workbook::new();
    
    let mut worksheet = Worksheet::new("Data".to_string());
    
    // Create header row
    for col in 1..=cols {
        let header = format!("Column{}", col);
        worksheet.set_cell_value(1, col, CellValue::from(header));
    }
    
    // Create data rows
    for row in 2..=rows + 1 {
        for col in 1..=cols {
            let value = match col {
                1 => CellValue::from(format!("Row{}", row - 1)),
                2 => CellValue::Number((row - 1) as f64),
                3 => CellValue::Number((row - 1) as f64 * 1.5),
                4 => CellValue::Boolean((row - 1) % 2 == 0),
                _ => CellValue::from(format!("Value{}-{}", row - 1, col)),
            };
            worksheet.set_cell_value(row, col, value);
        }
    }
    
    workbook.worksheets.push(worksheet);
    workbook.sheet_names.push("Data".to_string());
    
    workbook
}

fn benchmark_excel_creation(c: &mut Criterion) {
    let mut group = c.benchmark_group("excel_creation");
    group.sample_size(10); // Smaller sample for faster runs
    
    // Test different sizes: 1k, 10k, 50k, 100k rows
    let sizes = vec![1_000, 10_000, 50_000, 100_000];
    let cols = 12;
    
    for size in sizes {
        group.bench_with_input(
            BenchmarkId::new("create_workbook", size),
            &size,
            |b, &rows| {
                b.iter(|| {
                    black_box(create_large_workbook(rows, cols))
                });
            },
        );
        
        group.bench_with_input(
            BenchmarkId::new("create_and_save", size),
            &size,
            |b, &rows| {
                let temp_path = format!("/tmp/bench_{}.xlsx", rows);
                b.iter(|| {
                    let workbook = create_large_workbook(rows, cols);
                    workbook.save(&temp_path).unwrap();
                    // Clean up
                    let _ = std::fs::remove_file(&temp_path);
                });
            },
        );
    }
    
    group.finish();
}

fn benchmark_cell_operations(c: &mut Criterion) {
    let mut group = c.benchmark_group("cell_operations");
    
    group.bench_function("set_cell_string", |b| {
        let mut worksheet = Worksheet::new("Test".to_string());
        b.iter(|| {
            for row in 1..=1000 {
                for col in 1..=12 {
                    worksheet.set_cell_value(row, col, CellValue::from(format!("Value{}-{}", row, col)));
                }
            }
        });
    });
    
    group.bench_function("set_cell_number", |b| {
        let mut worksheet = Worksheet::new("Test".to_string());
        b.iter(|| {
            for row in 1..=1000 {
                for col in 1..=12 {
                    worksheet.set_cell_value(row, col, CellValue::Number((row * col) as f64));
                }
            }
        });
    });
    
    group.finish();
}

criterion_group!(benches, benchmark_excel_creation, benchmark_cell_operations);
criterion_main!(benches);
