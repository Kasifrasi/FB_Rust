//! Test: unprotect_range() mit verschiedenen Bereichsgrößen

use kmw_fb_rust::v2::Finanzbericht::header::write_header;
use kmw_fb_rust::v2::Finanzbericht::sheet_setup::sheet_setup;
use kmw_fb_rust::v2::Finanzbericht::styles::ReportStyles;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::Workbook;
use std::time::{Duration, Instant};
use std::fs;

fn test_unprotect_range(
    index: usize,
    last_row: u32,
    last_col: u16,
) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    sheet_setup(ws)?;
    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();
    
    // Entsperre großen Bereich mit einer Operation
    ws.unprotect_range(0, 0, last_row, last_col)?;

    let path = format!("/tmp/benchmark_large/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn benchmark_range(
    name: &str,
    last_row: u32,
    last_col: u16,
    num_files: usize,
) -> (Duration, Duration) {
    println!("\n📊 Test: {} (0:0 bis {}:{})", name, last_row, last_col);
    println!("   Generiere {} Dateien...", num_files);

    let total_start = Instant::now();
    let mut durations: Vec<Duration> = Vec::new();

    for i in 0..num_files {
        if let Ok(dur) = test_unprotect_range(i, last_row, last_col) {
            durations.push(dur);
        }
        if (i + 1) % 5 == 0 {
            print!(".");
        }
    }
    println!();

    let total = total_start.elapsed();
    let avg = durations.iter().sum::<Duration>() / durations.len() as u32;

    println!("   Gesamt: {:?} | Ø: {:?}", total, avg);

    (total, avg)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   BENCHMARK: unprotect_range() mit verschiedenen Größen      ║");
    println!("╚══════════════════════════════════════════════════════════════╝");

    let temp_dir = "/tmp/benchmark_large";
    fs::create_dir_all(temp_dir)?;

    const NUM_FILES: usize = 20;

    println!("\n🔥 TESTS MIT unprotect_range():");
    println!("═══════════════════════════════════════════════════════════════");

    let (total_1k_30, avg_1k_30) = benchmark_range("1000x30", 999, 29, NUM_FILES);
    let (total_10k_100, avg_10k_100) = benchmark_range("10000x100", 9999, 99, NUM_FILES);
    let (total_10k_10k, avg_10k_10k) = benchmark_range("10000x10000", 9999, 9999, NUM_FILES);

    println!("\n\n═══════════════════════════════════════════════════════════════");
    println!("                    ERGEBNISSE");
    println!("═══════════════════════════════════════════════════════════════");

    println!("\n📋 Vergleich:");
    println!("  1000x30        : Gesamt: {:?} | Ø: {:?}", total_1k_30, avg_1k_30);
    println!("  10000x100      : Gesamt: {:?} | Ø: {:?}", total_10k_100, avg_10k_100);
    println!("  10000x10000    : Gesamt: {:?} | Ø: {:?}", total_10k_10k, avg_10k_10k);

    let diff_100 = ((total_10k_100.as_secs_f64() - total_1k_30.as_secs_f64()) 
        / total_1k_30.as_secs_f64()) * 100.0;
    let diff_10k = ((total_10k_10k.as_secs_f64() - total_1k_30.as_secs_f64()) 
        / total_1k_30.as_secs_f64()) * 100.0;

    println!("\n💡 ANALYSE:");
    if diff_100.abs() < 10.0 {
        println!("   ✅ 10000x100 ist praktisch GLEICH schnell wie 1000x30!");
    } else if diff_100 > 0.0 {
        println!("   ⚠️  10000x100 ist {:.1}% langsamer", diff_100);
    } else {
        println!("   🚀 10000x100 ist {:.1}% SCHNELLER", diff_100.abs());
    }

    if diff_10k.abs() < 10.0 {
        println!("   ✅ 10000x10000 ist praktisch GLEICH schnell wie 1000x30!");
    } else if diff_10k > 0.0 {
        println!("   ⚠️  10000x10000 ist {:.1}% langsamer", diff_10k);
    } else {
        println!("   🚀 10000x10000 ist {:.1}% SCHNELLER", diff_10k.abs());
    }

    println!("\n📊 Durchsatz:");
    let throughput_1k = NUM_FILES as f64 / total_1k_30.as_secs_f64();
    let throughput_10k = NUM_FILES as f64 / total_10k_10k.as_secs_f64();
    println!("  1000x30     : {:.2} Dateien/Sekunde", throughput_1k);
    println!("  10000x10000 : {:.2} Dateien/Sekunde", throughput_10k);

    println!("\n🧹 Räume auf...");
    for i in 0..NUM_FILES {
        let _ = fs::remove_file(format!("{}/file_{:04}.xlsx", temp_dir, i));
    }
    let _ = fs::remove_dir(temp_dir);

    println!("✅ Fertig!");

    Ok(())
}
