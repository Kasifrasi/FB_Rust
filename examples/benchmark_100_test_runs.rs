//! Benchmark: Führe test_header_generation() 100 mal aus

use kmw_fb_rust::v2::Finanzbericht::header::write_header;
use kmw_fb_rust::v2::Finanzbericht::sheet_setup::sheet_setup;
use kmw_fb_rust::v2::Finanzbericht::styles::ReportStyles;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::time::{Duration, Instant};

fn test_header_generation(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();

    // 1. Add Target Sheet (Finanzbericht) - must be first to be on left
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;

    // 2. Build Translation Sheet (Reference for VLOOKUPs)
    build_trans_sheet(&mut workbook)?;

    // 3. Get Target Sheet back
    let ws = workbook
        .worksheet_from_name(sheet_name)
        .expect("Sheet not found");

    // 4. Set column format to unlocked for 1000 columns
    // This makes all cells in these columns unlocked by default
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for col in 0..1000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    // 5. Setup sheet (column widths, etc.)
    sheet_setup(ws)?;

    // 6. Prepare Styles
    let styles = ReportStyles::new();

    // 7. Write Header
    let suffix = "_de";
    let lang_val = "deutsch";
    write_header(ws, &styles, suffix, lang_val)?;

    // 8. Protect worksheet
    // All cells are unlocked by set_column_format()
    // Formulas are locked by write_formulas() which uses fmt.get_locked()
    ws.protect();

    // 9. Save to file for inspection
    let path = format!("/tmp/benchmark_100_runs/header_test_{:04}.xlsx", index);
    fs::create_dir_all("/tmp/benchmark_100_runs")?;
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   BENCHMARK: test_header_generation() 100 mal ausführen      ║");
    println!("╚══════════════════════════════════════════════════════════════╝\n");

    fs::create_dir_all("/tmp/benchmark_100_runs")?;

    const NUM_RUNS: usize = 100;

    println!("🔥 Starten 100 Test-Durchläufe...\n");
    println!("   Generiere {} Dateien...", NUM_RUNS);

    let total_start = Instant::now();
    let mut durations: Vec<Duration> = Vec::new();

    for i in 0..NUM_RUNS {
        if let Ok(dur) = test_header_generation(i) {
            durations.push(dur);
        }
        if (i + 1) % 10 == 0 {
            print!(".");
            std::io::Write::flush(&mut std::io::stdout())?;
        }
    }

    let total_duration = total_start.elapsed();
    println!("\n");

    // Berechne Statistiken
    durations.sort();
    let min = durations.first().unwrap();
    let max = durations.last().unwrap();
    let median = durations[durations.len() / 2];
    let avg = durations.iter().sum::<Duration>() / durations.len() as u32;

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

    println!("═══════════════════════════════════════════════════════════════");
    println!("                        ERGEBNISSE");
    println!("═══════════════════════════════════════════════════════════════\n");

    println!("📁 Anzahl Test-Durchläufe: {}", NUM_RUNS);
    println!();
    println!("⏱️  GESAMTZEIT:             {:?}", total_duration);
    println!();
    println!("📈 Pro Test-Durchlauf:");
    println!("   Minimum:               {:?}", min);
    println!("   Maximum:               {:?}", max);
    println!("   Median:                {:?}", median);
    println!("   Durchschnitt:          {:?}", avg);
    println!("   Standardabweichung:    {:.2} ms", std_dev);
    println!();
    println!(
        "🚀 Durchsatz:             {:.2} Tests/Sekunde",
        NUM_RUNS as f64 / total_duration.as_secs_f64()
    );

    println!("\n📊 ANALYSE:");
    if avg.as_millis() < 10 {
        println!("   ✅ Super schnell! ({:.2} ms pro Test)", avg.as_millis());
    } else if avg.as_millis() < 20 {
        println!("   ✅ Schnell! ({:.2} ms pro Test)", avg.as_millis());
    } else {
        println!("   ⚠️  Etwas langsam... ({:.2} ms pro Test)", avg.as_millis());
    }

    // Cleanup
    println!("\n🧹 Räume auf...");
    for entry in fs::read_dir("/tmp/benchmark_100_runs")? {
        let entry = entry?;
        let path = entry.path();
        if path.is_file() {
            fs::remove_file(&path)?;
        }
    }
    fs::remove_dir("/tmp/benchmark_100_runs")?;

    println!("✅ Fertig!");

    Ok(())
}
