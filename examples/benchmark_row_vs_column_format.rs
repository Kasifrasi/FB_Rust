//! Benchmark: set_row_format() vs set_column_format()
//! Vergleiche verschiedene Kombinationen

use kmw_fb_rust::v2::Finanzbericht::header::write_header;
use kmw_fb_rust::v2::Finanzbericht::sheet_setup::sheet_setup;
use kmw_fb_rust::v2::Finanzbericht::styles::ReportStyles;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::time::{Duration, Instant};

fn generate_with_10000_rows(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    // 10000 Zeilen unlocked
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for row in 0..10000u32 {
        ws.set_row_format(row, &unlocked)?;
    }

    sheet_setup(ws)?;
    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();

    let path = format!("/tmp/benchmark_rows/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn generate_with_1000_columns(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    // 1000 Spalten unlocked
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for col in 0..1000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    sheet_setup(ws)?;
    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();

    let path = format!("/tmp/benchmark_cols_1000/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn generate_with_1000_rows_10000_cols(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    // 1000 Zeilen + 10000 Spalten unlocked
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for row in 0..1000u32 {
        ws.set_row_format(row, &unlocked)?;
    }
    for col in 0..10000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    sheet_setup(ws)?;
    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();

    let path = format!("/tmp/benchmark_mixed/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn generate_with_10000_columns(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    // 10000 Spalten unlocked
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for col in 0..10000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    sheet_setup(ws)?;
    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();

    let path = format!("/tmp/benchmark_cols_10000/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn run_benchmark<F>(
    name: &str,
    generator: F,
    num_files: usize,
) -> (Duration, Duration)
where
    F: Fn(usize) -> Result<Duration, Box<dyn std::error::Error>>,
{
    println!("📊 Test: {}", name);
    println!("   Generiere {} Dateien...", num_files);
    
    let start = Instant::now();
    let mut durations = Vec::new();

    for i in 0..num_files {
        if let Ok(dur) = generator(i) {
            durations.push(dur);
        }
        if (i + 1) % 5 == 0 {
            print!(".");
        }
    }
    let total = start.elapsed();
    let avg = durations.iter().sum::<Duration>() / durations.len() as u32;
    println!("\n   Gesamt: {:?} | Ø: {:?}\n", total, avg);

    (total, avg)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   BENCHMARK: Zeilen vs Spalten Format                        ║");
    println!("╚══════════════════════════════════════════════════════════════╝\n");

    fs::create_dir_all("/tmp/benchmark_rows")?;
    fs::create_dir_all("/tmp/benchmark_cols_1000")?;
    fs::create_dir_all("/tmp/benchmark_mixed")?;
    fs::create_dir_all("/tmp/benchmark_cols_10000")?;

    const NUM_FILES: usize = 20;

    println!("🔥 TESTS:\n");

    let (total_10k_rows, avg_10k_rows) =
        run_benchmark("10000 Zeilen unlocked", generate_with_10000_rows, NUM_FILES);

    let (total_1k_cols, avg_1k_cols) =
        run_benchmark("1000 Spalten unlocked", generate_with_1000_columns, NUM_FILES);

    let (total_1k_rows_10k_cols, avg_1k_rows_10k_cols) =
        run_benchmark("1000 Zeilen + 10000 Spalten unlocked", generate_with_1000_rows_10000_cols, NUM_FILES);

    let (total_10k_cols, avg_10k_cols) =
        run_benchmark("10000 Spalten unlocked", generate_with_10000_columns, NUM_FILES);

    println!("═══════════════════════════════════════════════════════════════");
    println!("                    VERGLEICH");
    println!("═══════════════════════════════════════════════════════════════\n");

    println!("  10000 Zeilen          : Gesamt: {:?} | Ø: {:?}", total_10k_rows, avg_10k_rows);
    println!("  1000 Spalten          : Gesamt: {:?} | Ø: {:?}", total_1k_cols, avg_1k_cols);
    println!("  1000 Zeilen + 10000 Spalten: Gesamt: {:?} | Ø: {:?}", total_1k_rows_10k_cols, avg_1k_rows_10k_cols);
    println!("  10000 Spalten         : Gesamt: {:?} | Ø: {:?}", total_10k_cols, avg_10k_cols);

    println!("\n📊 SPEEDUP-Vergleiche:");
    let speedup_cols_vs_rows = total_10k_rows.as_secs_f64() / total_1k_cols.as_secs_f64();
    println!("  → 10000 Zeilen ist {:.1}x {} als 1000 Spalten",
        speedup_cols_vs_rows.max(1.0/speedup_cols_vs_rows),
        if speedup_cols_vs_rows > 1.0 { "langsamer" } else { "schneller" }
    );

    let speedup_10k_cols_vs_10k_rows = total_10k_rows.as_secs_f64() / total_10k_cols.as_secs_f64();
    println!("  → 10000 Zeilen ist {:.1}x {} als 10000 Spalten",
        speedup_10k_cols_vs_10k_rows.max(1.0/speedup_10k_cols_vs_10k_rows),
        if speedup_10k_cols_vs_10k_rows > 1.0 { "langsamer" } else { "schneller" }
    );

    let speedup_mixed = total_10k_rows.as_secs_f64() / total_1k_rows_10k_cols.as_secs_f64();
    println!("  → 10000 Zeilen ist {:.1}x {} als 1000 Zeilen + 10000 Spalten",
        speedup_mixed.max(1.0/speedup_mixed),
        if speedup_mixed > 1.0 { "langsamer" } else { "schneller" }
    );

    println!("\n💡 BEST PRACTICE:");
    let times = vec![
        ("10000 Zeilen", total_10k_rows),
        ("1000 Spalten", total_1k_cols),
        ("1000 Zeilen + 10000 Spalten", total_1k_rows_10k_cols),
        ("10000 Spalten", total_10k_cols),
    ];

    let (best_name, _) = times.iter().min_by_key(|(_, t)| t).unwrap();
    println!("   ✅ {} ist die schnellste Methode!", best_name);

    println!("\n🧹 Räume auf...");
    for dir in &["/tmp/benchmark_rows", "/tmp/benchmark_cols_1000", "/tmp/benchmark_mixed", "/tmp/benchmark_cols_10000"] {
        for i in 0..NUM_FILES {
            let _ = fs::remove_file(format!("{}/file_{:04}.xlsx", dir, i));
        }
        let _ = fs::remove_dir(dir);
    }

    println!("✅ Fertig!");

    Ok(())
}
