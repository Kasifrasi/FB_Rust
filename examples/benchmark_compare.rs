//! Vergleich verschiedener Konfigurationen und Threading-Strategien

use kmw_fb_rust::v2::Finanzbericht::header::{write_header, write_unlocked_base};
use kmw_fb_rust::v2::Finanzbericht::sheet_setup::sheet_setup;
use kmw_fb_rust::v2::Finanzbericht::styles::ReportStyles;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::Workbook;
use std::time::{Duration, Instant};
use std::fs;
use std::sync::{Arc, Mutex};
use std::thread;

fn generate_single_file(
    index: usize,
    rows: u32,
    cols: u16,
) -> Result<Duration, Box<dyn std::error::Error>> {
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;
    build_trans_sheet(&mut workbook)?;

    let ws = workbook.worksheet_from_name(sheet_name)?;
    write_unlocked_base(ws, rows, cols)?;
    sheet_setup(ws)?;

    let styles = ReportStyles::new();
    write_header(ws, &styles, "_de", "deutsch")?;
    ws.protect();

    let path = format!("/tmp/benchmark_compare/file_{:04}.xlsx", index);
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn benchmark_config(
    name: &str,
    rows: u32,
    cols: u16,
    num_files: usize,
) -> (Duration, Duration) {
    println!("\n📊 Test: {} ({} x {})", name, rows, cols);
    println!("   Generiere {} Dateien...", num_files);

    let total_start = Instant::now();
    let mut durations: Vec<Duration> = Vec::new();

    for i in 0..num_files {
        if let Ok(dur) = generate_single_file(i, rows, cols) {
            durations.push(dur);
        }
        if (i + 1) % 10 == 0 {
            print!(".");
        }
    }
    println!();

    let total = total_start.elapsed();
    let avg = durations.iter().sum::<Duration>() / durations.len() as u32;

    println!("   Gesamt: {:?} | Ø: {:?}", total, avg);

    (total, avg)
}

fn benchmark_config_threaded(
    name: &str,
    rows: u32,
    cols: u16,
    num_files: usize,
    num_threads: usize,
) -> (Duration, Duration) {
    println!("\n📊 Test (THREADED): {} ({} x {}) - {} Threads", name, rows, cols, num_threads);
    println!("   Generiere {} Dateien...", num_files);

    let total_start = Instant::now();
    let durations = Arc::new(Mutex::new(Vec::new()));

    let mut handles = vec![];

    let indices: Vec<usize> = (0..num_files).collect();
    let chunk_size = (num_files + num_threads - 1) / num_threads;

    for thread_id in 0..num_threads {
        let start_idx = thread_id * chunk_size;
        let end_idx = (start_idx + chunk_size).min(num_files);

        if start_idx >= num_files {
            break;
        }

        let durations_clone = durations.clone();
        let handle = thread::spawn(move || {
            let mut thread_durations = Vec::new();
            for i in start_idx..end_idx {
                if let Ok(dur) = generate_single_file(i, rows, cols) {
                    thread_durations.push(dur);
                }
                print!(".");
            }
            durations_clone.lock().unwrap().extend(thread_durations);
        });

        handles.push(handle);
    }

    for handle in handles {
        let _ = handle.join();
    }
    println!();

    let total = total_start.elapsed();
    let all_durations = durations.lock().unwrap();
    let avg = all_durations.iter().sum::<Duration>() / all_durations.len() as u32;

    println!("   Gesamt: {:?} | Ø: {:?}", total, avg);

    (total, avg)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   VERGLEICH: Konfigurationen & Threading                     ║");
    println!("╚══════════════════════════════════════════════════════════════╝");

    let temp_dir = "/tmp/benchmark_compare";
    fs::create_dir_all(temp_dir)?;

    const NUM_FILES: usize = 20;

    println!("\n🔥 SINGLE-THREADED TESTS:");
    println!("══════════════════════════════════════════════════════════════");

    let (total_1000_30, avg_1000_30) = benchmark_config("1000x30", 1000, 30, NUM_FILES);
    let (total_10000_100, avg_10000_100) = benchmark_config("10000x100", 10000, 100, NUM_FILES);

    println!("\n🔥 MULTI-THREADED TESTS (mit 1000x30):");
    println!("══════════════════════════════════════════════════════════════");

    let (total_threads_2, avg_threads_2) = benchmark_config_threaded("1000x30 MT", 1000, 30, NUM_FILES, 2);
    let (total_threads_4, avg_threads_4) = benchmark_config_threaded("1000x30 MT", 1000, 30, NUM_FILES, 4);
    let (total_threads_8, avg_threads_8) = benchmark_config_threaded("1000x30 MT", 1000, 30, NUM_FILES, 8);

    println!("\n\n═══════════════════════════════════════════════════════════════");
    println!("                    VERGLEICHSTABELLE");
    println!("═══════════════════════════════════════════════════════════════");

    println!("\n📋 Single-Threaded:");
    println!("  1000x30     : Gesamt: {:?} | Ø: {:?}", total_1000_30, avg_1000_30);
    println!("  10000x100   : Gesamt: {:?} | Ø: {:?}", total_10000_100, avg_10000_100);

    let diff_percent = ((total_10000_100.as_secs_f64() - total_1000_30.as_secs_f64()) 
        / total_1000_30.as_secs_f64()) * 100.0;
    println!("\n  → 10000x100 ist {:.1}% {}er", 
        diff_percent.abs(), 
        if diff_percent > 0.0 { "langsam" } else { "schnell" });

    println!("\n📋 Multi-Threaded (1000x30):");
    println!("  2 Threads   : Gesamt: {:?}", total_threads_2);
    println!("  4 Threads   : Gesamt: {:?}", total_threads_4);
    println!("  8 Threads   : Gesamt: {:?}", total_threads_8);

    let speedup_2 = total_1000_30.as_secs_f64() / total_threads_2.as_secs_f64();
    let speedup_4 = total_1000_30.as_secs_f64() / total_threads_4.as_secs_f64();
    let speedup_8 = total_1000_30.as_secs_f64() / total_threads_8.as_secs_f64();

    println!("\n  → Speedup mit 2 Threads: {:.2}x", speedup_2);
    println!("  → Speedup mit 4 Threads: {:.2}x", speedup_4);
    println!("  → Speedup mit 8 Threads: {:.2}x", speedup_8);

    println!("\n💡 ANALYSE:");
    if diff_percent > 0.0 {
        println!("   ⚠️  10000x100 ist LANGSAMER - lohnt sich also NICHT!");
    } else {
        println!("   ✅ 10000x100 ist SCHNELLER - könnte sich lohnen!");
    }

    if speedup_2 > 1.3 {
        println!("   ✅ Threading bringt SIGNIFIKANTE Speedups!");
    } else {
        println!("   ⚠️  Threading bringt WENIG Speedup - wahrscheinlich I/O-gebunden");
    }

    println!("\n🧹 Räume auf...");
    for i in 0..NUM_FILES {
        let _ = fs::remove_file(format!("{}/file_{:04}.xlsx", temp_dir, i));
    }
    let _ = fs::remove_dir(temp_dir);

    println!("✅ Fertig!");

    Ok(())
}
