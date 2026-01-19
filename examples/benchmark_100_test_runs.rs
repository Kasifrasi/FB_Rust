//! Performance-Benchmark: 100 Test-Durchläufe mit variabler Dateigröße
//!
//! Generiert multiple Dateien sequenziell und misst die Performance

use kmw_fb_rust::v2::report::formats::ReportStyles;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::values::ReportValues;
use kmw_fb_rust::v2::report::writer::write_report_v2;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::time::{Duration, Instant};

fn generate_file(
    index: usize,
    num_data_rows: usize,
) -> Result<Duration, Box<dyn std::error::Error>> {
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
        .with_project_number(&format!("TEST-{}", index))
        .with_project_title(&format!("Test {}", index))
        .with_exchange_rate(1.0);

    // Variiere die Datengröße
    use kmw_fb_rust::v2::report::cells::TableInputCell;
    let max_rows = std::cmp::min(num_data_rows, 5);
    for i in 0..max_rows as u8 {
        values
            .set(TableInputCell::ApprovedBudget(i), 1000.0 + i as f64 * 500.0)
            .set(
                TableInputCell::IncomeReportPeriod(i),
                100.0 + i as f64 * 200.0,
            )
            .set(TableInputCell::IncomeTotal(i), 500.0 + i as f64 * 300.0);
    }

    write_report_v2(ws, &styles, "_de", &values)?;
    ws.protect();

    let path = format!("/tmp/benchmark_runs/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║     BENCHMARK: 100 Test-Durchläufe mit Variationen           ║");
    println!("╚══════════════════════════════════════════════════════════════╝");
    println!();

    let temp_dir = "/tmp/benchmark_runs";
    fs::create_dir_all(temp_dir)?;

    println!("📊 Starte Tests...");
    println!();

    let mut durations: Vec<Duration> = Vec::new();
    let total_start = Instant::now();

    for i in 0..100 {
        let num_rows = (i % 5) + 1;
        let duration = generate_file(i, num_rows)?;
        durations.push(duration);

        if (i + 1) % 20 == 0 {
            println!("   {}/100 Tests abgeschlossen...", i + 1);
        }
    }

    let total_duration = total_start.elapsed();

    println!();
    println!("═══════════════════════════════════════════════════════════════");
    println!("                        ERGEBNISSE");
    println!("═══════════════════════════════════════════════════════════════");
    println!();

    let mut sorted = durations.clone();
    sorted.sort();
    let min = sorted[0];
    let max = sorted[99];
    let median = sorted[50];
    let avg = total_duration / 100u32;

    println!("📊 Tests:                 100");
    println!("⏱️  GESAMTZEIT:            {:?}", total_duration);
    println!();
    println!("📈 Pro Test:");
    println!("   Minimum:               {:?}", min);
    println!("   Maximum:               {:?}", max);
    println!("   Median:                {:?}", median);
    println!("   Durchschnitt:          {:?}", avg);
    println!();
    println!(
        "🚀 Durchsatz:             {:.2} Tests/Sekunde",
        100.0 / total_duration.as_secs_f64()
    );
    println!();

    println!("🧹 Räume auf...");
    for i in 0..100 {
        let path = format!("{}/file_{:04}.xlsx", temp_dir, i);
        let _ = fs::remove_file(&path);
    }
    let _ = fs::remove_dir(temp_dir);

    println!("✅ Fertig!");

    Ok(())
}
