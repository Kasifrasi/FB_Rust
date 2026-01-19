//! Vergleich: write_unlocked_base vs unprotect_range()

use kmw_fb_rust::v2::Finanzbericht::header::write_header;
use kmw_fb_rust::v2::Finanzbericht::sheet_setup::sheet_setup;
use kmw_fb_rust::v2::Finanzbericht::styles::ReportStyles;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::Workbook;
use std::time::{Duration, Instant};
use std::fs;

fn method_1_write_blanks(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    // Methode 1: Alle Zellen einzeln schreiben (OLD)
    let unlocked = rust_xlsxwriter::Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for row in 0..1000u32 {
        for col in 0..30u16 {
            ws.write_blank(row, col, &unlocked)?;
        }
    }

    sheet_setup(ws)?;
    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();

    let path = format!("/tmp/benchmark_unprotect/m1_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn method_2_unprotect_range(index: usize) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;

    // Methode 2: unprotect_range() (NEW - viel schneller!)
    sheet_setup(ws)?;
    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();
    ws.unprotect_range(0, 0, 999, 29)?;

    let path = format!("/tmp/benchmark_unprotect/m2_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   VERGLEICH: write_unlocked_base vs unprotect_range()        ║");
    println!("╚══════════════════════════════════════════════════════════════╝");

    let temp_dir = "/tmp/benchmark_unprotect";
    fs::create_dir_all(temp_dir)?;

    const NUM_FILES: usize = 20;

    println!("\n🔥 Methode 1: write_blank() für jede Zelle (30.000 writes)");
    println!("═══════════════════════════════════════════════════════════════");
    let mut durations_m1 = Vec::new();
    let start_m1 = Instant::now();
    for i in 0..NUM_FILES {
        let dur = method_1_write_blanks(i)?;
        durations_m1.push(dur);
        print!(".");
    }
    let total_m1 = start_m1.elapsed();
    let avg_m1 = durations_m1.iter().sum::<Duration>() / durations_m1.len() as u32;
    println!("\nGesamt: {:?} | Ø: {:?}", total_m1, avg_m1);

    println!("\n🚀 Methode 2: unprotect_range() (1 Operation)");
    println!("═══════════════════════════════════════════════════════════════");
    let mut durations_m2 = Vec::new();
    let start_m2 = Instant::now();
    for i in 0..NUM_FILES {
        let dur = method_2_unprotect_range(i)?;
        durations_m2.push(dur);
        print!(".");
    }
    let total_m2 = start_m2.elapsed();
    let avg_m2 = durations_m2.iter().sum::<Duration>() / durations_m2.len() as u32;
    println!("\nGesamt: {:?} | Ø: {:?}", total_m2, avg_m2);

    println!("\n\n═══════════════════════════════════════════════════════════════");
    println!("                       VERGLEICH");
    println!("═══════════════════════════════════════════════════════════════");

    let speedup = total_m1.as_secs_f64() / total_m2.as_secs_f64();
    let improvement_percent = ((total_m1.as_secs_f64() - total_m2.as_secs_f64()) 
        / total_m1.as_secs_f64()) * 100.0;

    println!("\nMethode 1 (write_blank):     {:?}", total_m1);
    println!("Methode 2 (unprotect_range): {:?}", total_m2);
    println!("\n🚀 Speedup:                  {:.2}x schneller", speedup);
    println!("💾 Verbesserung:             {:.1}% schneller", improvement_percent);

    println!("\n💡 FAZIT:");
    if speedup > 1.5 {
        println!("   ✅ unprotect_range() ist DEUTLICH schneller!");
        println!("   → Verwende diese Methode!");
    } else if speedup > 1.0 {
        println!("   ✅ unprotect_range() ist schneller");
        println!("   → Verwende diese Methode!");
    } else {
        println!("   ℹ️  Ähnliche Performance");
    }

    println!("\n🧹 Räume auf...");
    for i in 0..NUM_FILES {
        let _ = fs::remove_file(format!("{}/m1_{:04}.xlsx", temp_dir, i));
        let _ = fs::remove_file(format!("{}/m2_{:04}.xlsx", temp_dir, i));
    }
    let _ = fs::remove_dir(temp_dir);

    println!("✅ Fertig!");

    Ok(())
}
