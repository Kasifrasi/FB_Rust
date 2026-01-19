//! Benchmark: Führe test_header_generation() 100 mal aus mit Multi-Threading

use kmw_fb_rust::v2::Finanzbericht::header::write_header;
use kmw_fb_rust::v2::Finanzbericht::sheet_setup::sheet_setup;
use kmw_fb_rust::v2::Finanzbericht::styles::ReportStyles;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;
use std::sync::{Arc, Mutex};
use std::thread;
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
    ws.protect();

    // 9. Save to file for inspection
    let path = format!("/tmp/benchmark_100_runs_mt/header_test_{:04}.xlsx", index);
    fs::create_dir_all("/tmp/benchmark_100_runs_mt")?;
    workbook.save(&path)?;

    Ok(start.elapsed())
}

fn run_single_threaded(num_runs: usize) -> (Duration, Vec<Duration>) {
    println!("📊 Single-Threaded (1 Thread)");
    println!("   Generiere {} Dateien...", num_runs);

    let start = Instant::now();
    let mut durations: Vec<Duration> = Vec::new();

    for i in 0..num_runs {
        if let Ok(dur) = test_header_generation(i) {
            durations.push(dur);
        }
        if (i + 1) % 10 == 0 {
            print!(".");
            std::io::Write::flush(&mut std::io::stdout()).unwrap();
        }
    }

    let total = start.elapsed();
    println!();
    println!("   Gesamt: {:?}\n", total);

    (total, durations)
}

fn run_multi_threaded(num_runs: usize, num_threads: usize) -> (Duration, Vec<Duration>) {
    println!("📊 Multi-Threaded ({} Threads)", num_threads);
    println!("   Generiere {} Dateien...", num_runs);

    let durations = Arc::new(Mutex::new(Vec::new()));
    let mut handles = vec![];

    let chunk_size = (num_runs + num_threads - 1) / num_threads;
    let start = Instant::now();

    for thread_id in 0..num_threads {
        let start_idx = thread_id * chunk_size;
        let end_idx = (start_idx + chunk_size).min(num_runs);

        if start_idx >= num_runs {
            break;
        }

        let durations_clone = durations.clone();
        let handle = thread::spawn(move || {
            for i in start_idx..end_idx {
                if let Ok(dur) = test_header_generation(i) {
                    durations_clone.lock().unwrap().push(dur);
                }
                print!(".");
                std::io::Write::flush(&mut std::io::stdout()).unwrap();
            }
        });

        handles.push(handle);
    }

    for handle in handles {
        let _ = handle.join();
    }

    let total = start.elapsed();
    println!();
    println!("   Gesamt: {:?}\n", total);

    let final_durations = durations.lock().unwrap().clone();
    (total, final_durations)
}

fn calculate_stats(durations: &[Duration]) -> (Duration, Duration, Duration, f64) {
    let mut sorted = durations.to_vec();
    sorted.sort();

    let min = *sorted.first().unwrap();
    let max = *sorted.last().unwrap();
    let median = sorted[sorted.len() / 2];
    let avg = durations.iter().sum::<Duration>() / durations.len() as u32;

    let avg_nanos = avg.as_nanos() as f64;
    let variance: f64 = durations
        .iter()
        .map(|d| {
            let diff = d.as_nanos() as f64 - avg_nanos;
            diff * diff
        })
        .sum::<f64>()
        / durations.len() as f64;
    let std_dev = variance.sqrt() / 1_000_000.0;

    (min, max, median, std_dev)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║   BENCHMARK: test_header_generation() 100 mal Multi-Threaded ║");
    println!("╚══════════════════════════════════════════════════════════════╝\n");

    fs::create_dir_all("/tmp/benchmark_100_runs_mt")?;

    const NUM_RUNS: usize = 100;

    println!("🔥 SINGLE-THREADED TEST:\n");
    let (total_st, durations_st) = run_single_threaded(NUM_RUNS);
    let (min_st, max_st, median_st, std_st) = calculate_stats(&durations_st);
    let avg_st = durations_st.iter().sum::<Duration>() / durations_st.len() as u32;

    println!("🔥 MULTI-THREADED TESTS:\n");

    let (total_2, durations_2) = run_multi_threaded(NUM_RUNS, 2);
    let (min_2, _, median_2, std_2) = calculate_stats(&durations_2);
    let avg_2 = durations_2.iter().sum::<Duration>() / durations_2.len() as u32;

    let (total_4, durations_4) = run_multi_threaded(NUM_RUNS, 4);
    let (min_4, _, median_4, std_4) = calculate_stats(&durations_4);
    let avg_4 = durations_4.iter().sum::<Duration>() / durations_4.len() as u32;

    let (total_8, durations_8) = run_multi_threaded(NUM_RUNS, 8);
    let (min_8, _, median_8, std_8) = calculate_stats(&durations_8);
    let avg_8 = durations_8.iter().sum::<Duration>() / durations_8.len() as u32;

    println!("═══════════════════════════════════════════════════════════════");
    println!("                        ERGEBNISSE");
    println!("═══════════════════════════════════════════════════════════════\n");

    println!("📊 ZUSAMMENFASSUNG:");
    println!();
    println!("  Single-Thread:");
    println!("    Gesamt:      {:?}", total_st);
    println!("    Ø:           {:?}", avg_st);
    println!("    Min/Max:     {:?} / {:?}", min_st, max_st);
    println!("    Median:      {:?}", median_st);
    println!("    Std.abw.:    {:.2} ms", std_st);
    println!("    Durchsatz:   {:.2} Tests/Sec", NUM_RUNS as f64 / total_st.as_secs_f64());
    println!();

    println!("  2 Threads:");
    println!("    Gesamt:      {:?}", total_2);
    println!("    Ø:           {:?}", avg_2);
    println!("    Median:      {:?}", median_2);
    println!("    Std.abw.:    {:.2} ms", std_2);
    println!("    Durchsatz:   {:.2} Tests/Sec", NUM_RUNS as f64 / total_2.as_secs_f64());
    let speedup_2 = total_st.as_secs_f64() / total_2.as_secs_f64();
    println!("    Speedup:     {:.2}x", speedup_2);
    println!();

    println!("  4 Threads:");
    println!("    Gesamt:      {:?}", total_4);
    println!("    Ø:           {:?}", avg_4);
    println!("    Median:      {:?}", median_4);
    println!("    Std.abw.:    {:.2} ms", std_4);
    println!("    Durchsatz:   {:.2} Tests/Sec", NUM_RUNS as f64 / total_4.as_secs_f64());
    let speedup_4 = total_st.as_secs_f64() / total_4.as_secs_f64();
    println!("    Speedup:     {:.2}x", speedup_4);
    println!();

    println!("  8 Threads:");
    println!("    Gesamt:      {:?}", total_8);
    println!("    Ø:           {:?}", avg_8);
    println!("    Median:      {:?}", median_8);
    println!("    Std.abw.:    {:.2} ms", std_8);
    println!("    Durchsatz:   {:.2} Tests/Sec", NUM_RUNS as f64 / total_8.as_secs_f64());
    let speedup_8 = total_st.as_secs_f64() / total_8.as_secs_f64();
    println!("    Speedup:     {:.2}x", speedup_8);
    println!();

    println!("═══════════════════════════════════════════════════════════════");
    println!("                    ANALYSE");
    println!("═══════════════════════════════════════════════════════════════\n");

    println!("💡 SPEEDUP-ÜBERSICHT:");
    println!("  2 Threads:  {:.2}x schneller", speedup_2);
    println!("  4 Threads:  {:.2}x schneller", speedup_4);
    println!("  8 Threads:  {:.2}x schneller", speedup_8);

    if speedup_8 > 2.0 {
        println!("\n  ✅ Multi-Threading bringt DEUTLICHE Speedups!");
    } else if speedup_8 > 1.2 {
        println!("\n  ✅ Multi-Threading bringt moderate Speedups!");
    } else {
        println!("\n  ⚠️  Multi-Threading bringt wenig Speedup - wahrscheinlich I/O-gebunden");
    }

    // Cleanup
    println!("\n🧹 Räume auf...");
    for entry in fs::read_dir("/tmp/benchmark_100_runs_mt")? {
        let entry = entry?;
        let path = entry.path();
        if path.is_file() {
            fs::remove_file(&path)?;
        }
    }
    fs::remove_dir("/tmp/benchmark_100_runs_mt")?;

    println!("✅ Fertig!");

    Ok(())
}
