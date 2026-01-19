//! Benchmark: set_row_format() vs unprotect_range()

use kmw_fb_rust::v2::Finanzbericht::header::write_header;
use kmw_fb_rust::v2::Finanzbericht::sheet_setup::sheet_setup;
use kmw_fb_rust::v2::Finanzbericht::styles::ReportStyles;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::time::{Duration, Instant};

fn generate_with_set_row_format(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    // Methode 1: set_row_format() für 10000 Zeilen
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

    let path = format!("/tmp/benchmark_set_row/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn generate_with_unprotect_range(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
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
    ws.unprotect_range(0, 0, 9999, 9999)?;

    let path = format!("/tmp/benchmark_unprotect_comp/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   BENCHMARK: set_row_format() vs unprotect_range()          ║");
    println!("╚══════════════════════════════════════════════════════════════╝\n");

    let temp_dir1 = "/tmp/benchmark_set_row";
    let temp_dir2 = "/tmp/benchmark_unprotect_comp";
    fs::create_dir_all(temp_dir1)?;
    fs::create_dir_all(temp_dir2)?;

    const NUM_FILES: usize = 20;

    // Test 1: set_row_format()
    println!("📊 Test 1: set_row_format() für 10000 Zeilen");
    println!("   Generiere {} Dateien...", NUM_FILES);
    let start = Instant::now();
    let mut durations1: Vec<Duration> = Vec::new();

    for i in 0..NUM_FILES {
        let dur = generate_with_set_row_format(i)?;
        durations1.push(dur);
        if (i + 1) % 5 == 0 {
            print!(".");
        }
    }
    let total1 = start.elapsed();
    let avg1 = durations1.iter().sum::<Duration>() / durations1.len() as u32;
    println!("\n   Gesamt: {:?} | Ø: {:?}\n", total1, avg1);

    // Test 2: unprotect_range()
    println!("📊 Test 2: unprotect_range() 10000x10000");
    println!("   Generiere {} Dateien...", NUM_FILES);
    let start = Instant::now();
    let mut durations2: Vec<Duration> = Vec::new();

    for i in 0..NUM_FILES {
        let dur = generate_with_unprotect_range(i)?;
        durations2.push(dur);
        if (i + 1) % 5 == 0 {
            print!(".");
        }
    }
    let total2 = start.elapsed();
    let avg2 = durations2.iter().sum::<Duration>() / durations2.len() as u32;
    println!("\n   Gesamt: {:?} | Ø: {:?}\n", total2, avg2);

    // Ergebnisse
    println!("═══════════════════════════════════════════════════════════════");
    println!("                    VERGLEICH");
    println!("═══════════════════════════════════════════════════════════════\n");

    println!("  set_row_format() : Gesamt: {:?} | Ø: {:?}", total1, avg1);
    println!("  unprotect_range(): Gesamt: {:?} | Ø: {:?}", total2, avg2);

    let speedup = total1.as_secs_f64() / total2.as_secs_f64();
    println!("\n  → set_row_format() ist {:.1}x {} als unprotect_range()",
        speedup.max(1.0/speedup),
        if speedup > 1.0 { "langsamer" } else { "schneller" }
    );

    println!("\n💡 ANALYSE:");
    if speedup > 1.5 {
        println!("   ⚠️  set_row_format() ist DEUTLICH langsamer - nicht empfohlen!");
    } else if speedup > 1.1 {
        println!("   ⚠️  set_row_format() ist etwas langsamer, aber akzeptabel");
    } else {
        println!("   ✅ set_row_format() ist praktisch gleich schnell!");
    }

    // Cleanup
    println!("\n🧹 Räume auf...");
    for i in 0..NUM_FILES {
        let _ = fs::remove_file(format!("{}/file_{:04}.xlsx", temp_dir1, i));
        let _ = fs::remove_file(format!("{}/file_{:04}.xlsx", temp_dir2, i));
    }
    let _ = fs::remove_dir(temp_dir1);
    let _ = fs::remove_dir(temp_dir2);

    println!("✅ Fertig!");

    Ok(())
}
