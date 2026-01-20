//! Benchmark: Finanzbericht-Generierung
//!
//! Vergleicht single-threaded vs multi-threaded Performance
//! bei 100 und 1000 Dateien mit unterschiedlichen Parametern.
//!
//! Usage:
//!   cargo run --example benchmark --release
//!
//! Die generierten Dateien werden in einem temporären Verzeichnis erstellt
//! und nach dem Benchmark gelöscht.

use kmw_fb_rust::v2::common::{LANG_CONFIG, LANG_SUFFIXES};
use kmw_fb_rust::v2::lang::build_sheet as build_sprachversionen;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::ApiKey;
use kmw_fb_rust::v2::report::{write_report_v2_with_body, BodyConfig, ReportStyles, ReportValues};
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::path::Path;
use std::sync::Arc;
use std::thread;
use std::time::Instant;

/// Sprachen für den Benchmark
const LANGUAGES: [&str; 5] = [
    "Deutsch",
    "Englisch",
    "Französisch",
    "Spanisch",
    "Portugiesisch",
];

/// Generiert einen einzelnen Finanzbericht mit variablen Parametern
fn generate_report(
    index: usize,
    output_dir: &Path,
) -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    // Sprache basierend auf Index rotieren
    let lang_key = LANGUAGES[index % LANGUAGES.len()];
    let config = LANG_CONFIG.get(lang_key).unwrap();
    let suffix = LANG_SUFFIXES.get(lang_key).unwrap();

    let mut workbook = Workbook::new();
    let styles = ReportStyles::new();

    // Sprachversionen-Sheet
    build_sprachversionen(&mut workbook)?;

    // Worksheet erstellen
    let ws = workbook.add_worksheet();
    ws.set_name(config.fb_sheet)?;

    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();
    for col in 0..1000u16 {
        ws.set_column_format(col, &unlocked).ok();
    }

    setup_sheet(ws)?;

    // Variable Werte basierend auf Index
    let mut values = ReportValues::new();
    values.set(ApiKey::Language, config.lang_val);
    values.set(ApiKey::Currency, if index % 2 == 0 { "EUR" } else { "USD" });
    values.set(ApiKey::ProjectNumber, format!("PROJ-{:05}", index));
    values.set(
        ApiKey::ProjectTitle,
        format!("Project {} ({})", index, config.lang_val),
    );
    values.set(
        ApiKey::ProjectStart,
        format!("01.{:02}.2024", (index % 12) + 1),
    );
    values.set(ApiKey::ProjectEnd, "31.12.2024");
    values.set(
        ApiKey::ReportStart,
        format!("01.{:02}.2024", (index % 12) + 1),
    );
    values.set(ApiKey::ReportEnd, "30.06.2024");

    // Variable Einnahmen
    let base_budget = 10000.0 + (index as f64 * 100.0);
    for i in 0..5u8 {
        values.set(ApiKey::ApprovedBudget(i), base_budget * (i + 1) as f64);
        values.set(
            ApiKey::IncomeReportPeriod(i),
            base_budget * 0.5 * (i + 1) as f64,
        );
        values.set(ApiKey::IncomeTotal(i), base_budget * 0.5 * (i + 1) as f64);
    }

    // Variable Body-Konfiguration
    let pos_offset = (index % 5) as u16;
    let body_config = BodyConfig::new()
        .with_positions(1, 5 + pos_offset)
        .with_positions(2, 3 + pos_offset)
        .with_positions(3, 4 + pos_offset)
        .with_positions(4, 3 + pos_offset)
        .with_positions(5, 2 + pos_offset)
        .with_positions(6, 0)
        .with_positions(7, 0)
        .with_positions(8, 0);

    // Kostenpositionen mit variablen Werten
    let base_cost = 1000.0 + (index as f64 * 10.0);
    for cat in 1..=5u8 {
        let num_pos = match cat {
            1 => 5 + pos_offset,
            2 => 3 + pos_offset,
            3 => 4 + pos_offset,
            4 => 3 + pos_offset,
            5 => 2 + pos_offset,
            _ => 0,
        };
        for pos in 1..=num_pos {
            let cost = base_cost * (cat as f64) * (pos as f64 / 2.0);
            values.set_position_row(
                cat,
                pos,
                format!("Position {}.{}", cat, pos),
                cost,
                cost * 0.5,
                cost * 0.5,
                "",
            );
        }
    }

    // Header-Eingabe Kategorien
    values.set_header_input(6, base_cost * 4.0, base_cost * 2.0, base_cost * 2.0, "");
    values.set_header_input(7, base_cost * 6.0, base_cost * 3.0, base_cost * 3.0, "");
    values.set_header_input(8, base_cost * 2.5, base_cost * 1.25, base_cost * 1.25, "");

    // Footer-Werte
    values.set_footer_bank(base_cost * 10.0);
    values.set_footer_kasse(base_cost * 2.0);
    values.set_footer_sonstiges(base_cost * 0.5);

    // Report schreiben
    write_report_v2_with_body(ws, &styles, suffix, &values, &body_config)?;
    ws.protect();

    // Speichern
    let filename = output_dir.join(format!("report_{:05}.xlsx", index));
    workbook.save(&filename)?;

    Ok(())
}

/// Benchmark single-threaded
fn benchmark_single_threaded(count: usize, output_dir: &Path) -> std::time::Duration {
    let start = Instant::now();

    for i in 0..count {
        if let Err(e) = generate_report(i, output_dir) {
            eprintln!("Error generating report {}: {}", i, e);
        }
    }

    start.elapsed()
}

/// Benchmark multi-threaded
fn benchmark_multi_threaded(
    count: usize,
    output_dir: &Path,
    num_threads: usize,
) -> std::time::Duration {
    let start = Instant::now();
    let output_dir = Arc::new(output_dir.to_path_buf());

    let chunk_size = (count + num_threads - 1) / num_threads;
    let mut handles = Vec::new();

    for thread_id in 0..num_threads {
        let output_dir = Arc::clone(&output_dir);
        let start_idx = thread_id * chunk_size;
        let end_idx = std::cmp::min(start_idx + chunk_size, count);

        let handle = thread::spawn(move || {
            for i in start_idx..end_idx {
                if let Err(e) = generate_report(i, &output_dir) {
                    eprintln!("Error generating report {}: {}", i, e);
                }
            }
        });

        handles.push(handle);
    }

    for handle in handles {
        handle.join().unwrap();
    }

    start.elapsed()
}

fn run_benchmark(count: usize, output_dir: &Path) {
    println!("\n{}", "=".repeat(60));
    println!("  BENCHMARK: {} Dateien", count);
    println!("{}", "=".repeat(60));

    // Verzeichnis leeren
    if output_dir.exists() {
        fs::remove_dir_all(output_dir).ok();
    }
    fs::create_dir_all(output_dir).unwrap();

    // Single-threaded
    print!("\n  Single-threaded... ");
    std::io::Write::flush(&mut std::io::stdout()).unwrap();
    let single_time = benchmark_single_threaded(count, output_dir);
    println!("{:.2?}", single_time);
    println!(
        "    -> {:.2} Dateien/Sekunde",
        count as f64 / single_time.as_secs_f64()
    );

    // Cleanup
    fs::remove_dir_all(output_dir).ok();
    fs::create_dir_all(output_dir).unwrap();

    // Multi-threaded mit verschiedenen Thread-Anzahlen
    let thread_counts = [2, 4, 8, 16];

    for &num_threads in &thread_counts {
        print!("\n  Multi-threaded ({} threads)... ", num_threads);
        std::io::Write::flush(&mut std::io::stdout()).unwrap();
        let multi_time = benchmark_multi_threaded(count, output_dir, num_threads);
        println!("{:.2?}", multi_time);
        println!(
            "    -> {:.2} Dateien/Sekunde",
            count as f64 / multi_time.as_secs_f64()
        );
        println!(
            "    -> Speedup: {:.2}x",
            single_time.as_secs_f64() / multi_time.as_secs_f64()
        );

        // Cleanup zwischen Runs
        fs::remove_dir_all(output_dir).ok();
        fs::create_dir_all(output_dir).unwrap();
    }

    // Finale Cleanup
    fs::remove_dir_all(output_dir).ok();
}

fn main() {
    println!("\n{}", "#".repeat(60));
    println!("  FINANZBERICHT BENCHMARK");
    println!("  Vergleich: Single-threaded vs Multi-threaded");
    println!("{}", "#".repeat(60));

    let cpu_count = thread::available_parallelism()
        .map(|p| p.get())
        .unwrap_or(4);
    println!("\n  Verfügbare CPU-Kerne: {}", cpu_count);

    let temp_dir = std::env::temp_dir().join("finanzbericht_benchmark");

    // Benchmark 100 Dateien
    run_benchmark(100, &temp_dir);

    // Benchmark 1000 Dateien
    run_benchmark(1000, &temp_dir);

    println!("\n{}", "#".repeat(60));
    println!("  BENCHMARK ABGESCHLOSSEN");
    println!("{}\n", "#".repeat(60));
}
