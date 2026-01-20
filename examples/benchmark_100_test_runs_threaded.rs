//! Performance-Benchmark: 100 Test-Durchläufe mit Threading
//!
//! Vergleicht sequenzielle vs. parallele Verarbeitung

use kmw_fb_rust::v2::report::api::ApiKey;
use kmw_fb_rust::v2::report::formats::ReportStyles;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::values::ReportValues;
use kmw_fb_rust::v2::report::writer::write_report_v2;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::thread;
use std::time::{Duration, Instant};

fn generate_file_sequential(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;

    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for col in 0..1000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    setup_sheet(ws)?;
    let styles = ReportStyles::new();

    let mut values = ReportValues::new()
        .with_language("deutsch")
        .with_currency("EUR")
        .with_project_number(&format!("SEQ-{:03}", index))
        .with_project_title(&format!("Sequential Test {}", index));

    for i in 0..5u8 {
        values
            .set(ApiKey::ApprovedBudget(i), 1000.0 + i as f64 * 500.0)
            .set(ApiKey::IncomeReportPeriod(i), 100.0 + i as f64 * 200.0)
            .set(ApiKey::IncomeTotal(i), 500.0 + i as f64 * 300.0);
    }

    write_report_v2(ws, &styles, "_de", &values)?;
    ws.protect();

    let path = format!("/tmp/benchmark_threaded/seq_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn generate_file_threaded(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;

    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for col in 0..1000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    setup_sheet(ws)?;
    let styles = ReportStyles::new();

    let mut values = ReportValues::new()
        .with_language("deutsch")
        .with_currency("EUR")
        .with_project_number(&format!("THR-{:03}", index))
        .with_project_title(&format!("Threaded Test {}", index));

    for i in 0..5u8 {
        values
            .set(ApiKey::ApprovedBudget(i), 1000.0 + i as f64 * 500.0)
            .set(ApiKey::IncomeReportPeriod(i), 100.0 + i as f64 * 200.0)
            .set(ApiKey::IncomeTotal(i), 500.0 + i as f64 * 300.0);
    }

    write_report_v2(ws, &styles, "_de", &values)?;
    ws.protect();

    let path = format!("/tmp/benchmark_threaded/thr_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn run_sequential(count: usize) -> Duration {
    let total_start = Instant::now();
    for i in 0..count {
        let _ = generate_file_sequential(i);
    }
    total_start.elapsed()
}

fn run_threaded(count: usize, num_threads: usize) -> Duration {
    let total_start = Instant::now();
    let indices: Vec<usize> = (0..count).collect();
    let chunk_size = (count + num_threads - 1) / num_threads;

    let mut handles = vec![];

    for chunk in indices.chunks(chunk_size) {
        let chunk = chunk.to_vec();
        let handle = thread::spawn(move || {
            for idx in chunk {
                let _ = generate_file_threaded(idx);
            }
        });
        handles.push(handle);
    }

    for handle in handles {
        let _ = handle.join();
    }

    total_start.elapsed()
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║  BENCHMARK: Sequential vs. Threaded (100 Dateien)            ║");
    println!("╚══════════════════════════════════════════════════════════════╝");
    println!();

    let temp_dir = "/tmp/benchmark_threaded";
    fs::create_dir_all(temp_dir)?;

    const COUNT: usize = 100;

    // Sequenzielle Verarbeitung
    println!("🔄 Starte sequenzielle Verarbeitung (1 Thread)...");
    let seq_duration = run_sequential(COUNT);
    println!("   ✓ Fertig in {:?}", seq_duration);
    println!();

    // Threading mit 2 Threads
    println!("🔄 Starte 2-Thread Verarbeitung...");
    let thr2_duration = run_threaded(COUNT, 2);
    println!("   ✓ Fertig in {:?}", thr2_duration);
    println!();

    // Threading mit 4 Threads
    println!("🔄 Starte 4-Thread Verarbeitung...");
    let thr4_duration = run_threaded(COUNT, 4);
    println!("   ✓ Fertig in {:?}", thr4_duration);
    println!();

    // Threading mit 8 Threads
    println!("🔄 Starte 8-Thread Verarbeitung...");
    let thr8_duration = run_threaded(COUNT, 8);
    println!("   ✓ Fertig in {:?}", thr8_duration);
    println!();

    println!("═══════════════════════════════════════════════════════════════");
    println!("                        ERGEBNISSE");
    println!("═══════════════════════════════════════════════════════════════");
    println!();

    println!("📁 Anzahl Dateien:        {}", COUNT);
    println!();
    println!("⏱️  ZEITEN:");
    println!("   1 Thread (seq):        {:?}", seq_duration);
    println!("   2 Threads:             {:?}", thr2_duration);
    println!("   4 Threads:             {:?}", thr4_duration);
    println!("   8 Threads:             {:?}", thr8_duration);
    println!();
    println!("📈 DURCHSATZ:");
    println!(
        "   1 Thread (seq):        {:.2} Dateien/Sekunde",
        COUNT as f64 / seq_duration.as_secs_f64()
    );
    println!(
        "   2 Threads:             {:.2} Dateien/Sekunde",
        COUNT as f64 / thr2_duration.as_secs_f64()
    );
    println!(
        "   4 Threads:             {:.2} Dateien/Sekunde",
        COUNT as f64 / thr4_duration.as_secs_f64()
    );
    println!(
        "   8 Threads:             {:.2} Dateien/Sekunde",
        COUNT as f64 / thr8_duration.as_secs_f64()
    );
    println!();

    // Speedup Berechnung
    let speedup_2 = seq_duration.as_secs_f64() / thr2_duration.as_secs_f64();
    let speedup_4 = seq_duration.as_secs_f64() / thr4_duration.as_secs_f64();
    let speedup_8 = seq_duration.as_secs_f64() / thr8_duration.as_secs_f64();

    println!("🚀 SPEEDUP vs. Sequential:");
    println!("   2 Threads:             {:.2}x schneller", speedup_2);
    println!("   4 Threads:             {:.2}x schneller", speedup_4);
    println!("   8 Threads:             {:.2}x schneller", speedup_8);
    println!();

    println!("🧹 Räume auf...");
    for i in 0..COUNT {
        let _ = fs::remove_file(format!("/tmp/benchmark_threaded/seq_{:04}.xlsx", i));
        let _ = fs::remove_file(format!("/tmp/benchmark_threaded/thr_{:04}.xlsx", i));
    }
    let _ = fs::remove_dir(temp_dir);

    println!("✅ Fertig!");

    Ok(())
}
