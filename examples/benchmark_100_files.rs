//! Performance-Benchmark: 100 Excel-Dateien mit Testdaten generieren
//!
//! Nutzt die neue V2 API mit write_report_v2()
//! Misst die Generierungszeit mit vollständigen API-Werten

use kmw_fb_rust::v2::report::api::ApiKey;
use kmw_fb_rust::v2::report::formats::ReportStyles;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::values::ReportValues;
use kmw_fb_rust::v2::report::writer::write_report_v2;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::time::{Duration, Instant};

const NUM_FILES: usize = 100;
const WARMUP_RUNS: usize = 3;

fn generate_single_file(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();

    // 1. Add Target Sheet
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;

    // 2. Build Translation Sheet
    build_trans_sheet(&mut workbook)?;

    // 3. Get Target Sheet back
    let ws = workbook.worksheet_from_name(sheet_name)?;

    // 4. Set column format to unlocked for all columns
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for col in 0..1000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    // 5. Setup sheet
    setup_sheet(ws)?;

    // 6. Prepare Styles
    let styles = ReportStyles::new();

    // 7. Prepare Values mit Testdaten
    let mut values = ReportValues::new()
        .with_language("deutsch")
        .with_currency("EUR")
        .with_project_number(&format!("PROJ-{:05}", index))
        .with_project_title(&format!("Test Projekt Nr. {}", index))
        .with_project_start("2024-01-01")
        .with_project_end("2024-12-31")
        .with_report_start("2024-01-01")
        .with_report_end("2024-03-31");

    // Tabellendaten setzen (D15-F19)
    for i in 0..5u8 {
        let budget = 1000.0 + i as f64 * 500.0;
        let income_period = 100.0 + i as f64 * 200.0;
        let income_total = 500.0 + i as f64 * 300.0;

        values
            .set(ApiKey::ApprovedBudget(i), budget)
            .set(ApiKey::IncomeReportPeriod(i), income_period)
            .set(ApiKey::IncomeTotal(i), income_total);
    }

    // 8. Write Report mit V2 Writer
    let suffix = "_de";
    write_report_v2(ws, &styles, suffix, &values)?;

    // 9. Protect worksheet
    ws.protect();

    // 10. Save to file
    let path = format!("/tmp/benchmark_test/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn calculate_statistics(durations: &[Duration]) -> (Duration, Duration, Duration, Duration, f64) {
    let mut sorted: Vec<Duration> = durations.to_vec();
    sorted.sort();

    let min = *sorted.first().unwrap();
    let max = *sorted.last().unwrap();
    let median = sorted[sorted.len() / 2];

    let total: Duration = durations.iter().sum();
    let avg = total / durations.len() as u32;

    // Standardabweichung
    let avg_nanos = avg.as_nanos() as f64;
    let variance: f64 = durations
        .iter()
        .map(|d| {
            let diff = d.as_nanos() as f64 - avg_nanos;
            diff * diff
        })
        .sum::<f64>()
        / durations.len() as f64;
    let std_dev = variance.sqrt() / 1_000_000.0; // in ms

    (min, max, median, avg, std_dev)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   BENCHMARK: 100 Excel-Dateien mit Testdaten generieren      ║");
    println!("╚══════════════════════════════════════════════════════════════╝");
    println!();

    // Erstelle temp Verzeichnis
    let temp_dir = "/tmp/benchmark_test";
    fs::create_dir_all(temp_dir)?;

    // Warmup
    println!("🔥 Warmup ({} Durchläufe)...", WARMUP_RUNS);
    for i in 0..WARMUP_RUNS {
        let duration = generate_single_file(i)?;
        println!("   Warmup {}: {:?}", i + 1, duration);
    }
    println!();

    // Cleanup nach Warmup
    for i in 0..WARMUP_RUNS {
        let path = format!("{}/file_{:04}.xlsx", temp_dir, i);
        let _ = fs::remove_file(&path);
    }

    // Hauptmessung
    println!("📊 Starte Benchmark mit {} Dateien...", NUM_FILES);
    println!();

    let mut durations: Vec<Duration> = Vec::with_capacity(NUM_FILES);
    let total_start = Instant::now();

    for i in 0..NUM_FILES {
        let duration = generate_single_file(i)?;
        durations.push(duration);

        // Fortschritt alle 10 Dateien
        if (i + 1) % 10 == 0 {
            println!("   {} / 100 Dateien generiert...", i + 1);
        }
    }

    let total_duration = total_start.elapsed();

    println!();
    println!("═══════════════════════════════════════════════════════════════");
    println!("                        ERGEBNISSE");
    println!("═══════════════════════════════════════════════════════════════");
    println!();

    let (min, max, median, avg, std_dev) = calculate_statistics(&durations);

    println!("📁 Anzahl Dateien:        {}", NUM_FILES);
    println!();
    println!("⏱️  GESAMTZEIT:            {:?}", total_duration);
    println!();
    println!("📈 Pro Datei:");
    println!("   Minimum:               {:?}", min);
    println!("   Maximum:               {:?}", max);
    println!("   Median:                {:?}", median);
    println!("   Durchschnitt:          {:?}", avg);
    println!("   Standardabweichung:    {:.2} ms", std_dev);
    println!();
    println!(
        "🚀 Durchsatz:             {:.2} Dateien/Sekunde",
        NUM_FILES as f64 / total_duration.as_secs_f64()
    );
    println!();

    // Cleanup
    println!("🧹 Räume auf...");
    for i in 0..NUM_FILES {
        let path = format!("{}/file_{:04}.xlsx", temp_dir, i);
        let _ = fs::remove_file(&path);
    }
    let _ = fs::remove_dir(temp_dir);

    println!("✅ Fertig!");

    Ok(())
}
