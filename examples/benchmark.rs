//! Benchmark: Finanzbericht-Generierung
//!
//! Professioneller Throughput-Benchmark mit Warmup, statistischer Auswertung
//! und automatischem README-Update.
//!
//! Usage:
//!   cargo run --example benchmark --release
//!
//! Die Ergebnisse werden auf der Konsole ausgegeben und in README.md geschrieben
//! (zwischen <!-- PERF_START --> und <!-- PERF_END --> Markern).

use kmw_fb_rust::{precompute_hash_with_spin_count, PositionEntry, ReportConfig, TableEntry};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant};
use sysinfo::System;

// ============================================================================
// Konfiguration
// ============================================================================

const LANGUAGES: [&str; 5] = ["deutsch", "english", "francais", "espanol", "portugues"];
const CURRENCIES: [&str; 5] = ["EUR", "USD", "GBP", "CHF", "BRL"];

/// Benchmark-Konfigurationen: (file_count, warmup, samples, multi_only)
const BENCH_CONFIGS: &[(usize, usize, usize, bool)] = &[
    (100, 5, 20, false),
    (1_000, 3, 10, false),
    (10_000, 1, 3, true),
];

/// Thread-Anzahlen für Multi-threaded-Benchmarks
const THREAD_COUNTS: &[usize] = &[2, 4, 8, 16];

/// Thread-Anzahl für Protected-Benchmarks
const PROT_THREADS: usize = 8;

// ============================================================================
// Hardware-Info
// ============================================================================

struct HardwareInfo {
    cpu: String,
    physical_cores: usize,
    logical_cores: usize,
    ram_gb: u64,
    os: String,
}

impl HardwareInfo {
    fn collect() -> Self {
        let mut sys = System::new();
        sys.refresh_cpu_all();
        sys.refresh_memory();

        let cpu = sys
            .cpus()
            .first()
            .map(|c| c.brand().to_string())
            .unwrap_or_else(|| "Unknown CPU".to_string());

        let logical_cores = sys.cpus().len();
        let physical_cores = sys.physical_core_count().unwrap_or(logical_cores);
        let ram_gb = sys.total_memory() / 1_073_741_824;
        let os = System::long_os_version().unwrap_or_else(|| "Unknown OS".to_string());

        Self {
            cpu,
            physical_cores,
            logical_cores,
            ram_gb,
            os,
        }
    }

    fn summary(&self) -> String {
        format!(
            "{} · {} cores ({} logical) · {} GB RAM · {}",
            self.cpu, self.physical_cores, self.logical_cores, self.ram_gb, self.os
        )
    }
}

// ============================================================================
// Statistik
// ============================================================================

struct BenchStats {
    /// Throughput-Samples: Dateien/Sekunde pro Messung
    samples: Vec<f64>,
}

impl BenchStats {
    fn new(samples: Vec<f64>) -> Self {
        assert!(!samples.is_empty(), "BenchStats requires at least one sample");
        Self { samples }
    }

    fn mean(&self) -> f64 {
        self.samples.iter().sum::<f64>() / self.samples.len() as f64
    }

    fn std_dev(&self) -> f64 {
        if self.samples.len() < 2 {
            return 0.0;
        }
        let mean = self.mean();
        let variance = self.samples.iter().map(|x| (x - mean).powi(2)).sum::<f64>()
            / (self.samples.len() - 1) as f64;
        variance.sqrt()
    }

    fn cv_percent(&self) -> f64 {
        let mean = self.mean();
        if mean == 0.0 {
            return 0.0;
        }
        self.std_dev() / mean * 100.0
    }

    fn min(&self) -> f64 {
        self.samples.iter().cloned().fold(f64::INFINITY, f64::min)
    }

    fn max(&self) -> f64 {
        self.samples.iter().cloned().fold(f64::NEG_INFINITY, f64::max)
    }

    fn median(&self) -> f64 {
        let mut sorted = self.samples.clone();
        sorted.sort_by(|a, b| a.partial_cmp(b).unwrap());
        let n = sorted.len();
        if n % 2 == 0 {
            (sorted[n / 2 - 1] + sorted[n / 2]) / 2.0
        } else {
            sorted[n / 2]
        }
    }

    /// Mittlere Dauer pro Batch (aus Throughput berechnet)
    fn mean_duration(&self, file_count: usize) -> Duration {
        let mean_throughput = self.mean();
        if mean_throughput == 0.0 {
            return Duration::ZERO;
        }
        Duration::from_secs_f64(file_count as f64 / mean_throughput)
    }

    /// Qualitätsurteil basierend auf CV
    fn quality(&self) -> &'static str {
        match self.cv_percent() as u32 {
            0..=4 => "excellent",
            5..=9 => "good",
            10..=19 => "fair",
            _ => "noisy",
        }
    }
}

// ============================================================================
// Benchmark-Ergebnis
// ============================================================================

struct BenchResult {
    label: String,
    file_count: usize,
    threads: usize,
    stats: BenchStats,
}

impl BenchResult {
    fn throughput_mean(&self) -> f64 {
        self.stats.mean()
    }
}

// ============================================================================
// Report-Generierung
// ============================================================================

fn build_config(index: usize) -> ReportConfig {
    let lang = LANGUAGES[index % 5];
    let pos_offset = (index % 5) as u16;
    let base_cost = 1000.0 + (index as f64 * 10.0);
    let base_budget = 10_000.0 + (index as f64 * 100.0);

    let table: Vec<TableEntry> = (0..5u8)
        .map(|i| TableEntry {
            index: i,
            approved_budget: Some(base_budget * (i + 1) as f64),
            income_report: Some(base_budget * 0.5 * (i + 1) as f64),
            income_total: Some(base_budget * 0.5 * (i + 1) as f64),
            reason: None,
        })
        .collect();

    let cat_counts: [(u8, u16); 8] = [
        (1, 5 + pos_offset),
        (2, 3 + pos_offset),
        (3, 4 + pos_offset),
        (4, 3 + pos_offset),
        (5, 2 + pos_offset),
        (6, 0),
        (7, 0),
        (8, 0),
    ];

    let mut positions = Vec::new();
    for &(cat, count) in &cat_counts {
        if count == 0 {
            let (a, b, c) = match cat {
                6 => (4.0, 2.0, 2.0),
                7 => (6.0, 3.0, 3.0),
                _ => (2.5, 1.25, 1.25),
            };
            positions.push(PositionEntry {
                category: cat,
                position: 0,
                description: None,
                approved: Some(base_cost * a),
                income_report: Some(base_cost * b),
                income_total: Some(base_cost * c),
                remark: None,
            });
        } else {
            for pos in 1..=count {
                let cost = base_cost * (cat as f64) * (pos as f64 / 2.0);
                positions.push(PositionEntry {
                    category: cat,
                    position: pos,
                    description: Some(format!("Position {}.{}", cat, pos)),
                    approved: Some(cost),
                    income_report: Some(cost * 0.5),
                    income_total: Some(cost * 0.5),
                    remark: None,
                });
            }
        }
    }

    ReportConfig {
        language: lang.to_string(),
        currency: CURRENCIES[index % 5].to_string(),
        project_number: Some(format!("PROJ-{:05}", index)),
        project_title: Some(format!("Project {} ({})", index, lang)),
        project_start: Some(format!("01.{:02}.2024", (index % 12) + 1)),
        project_end: Some("31.12.2024".to_string()),
        report_start: Some(format!("01.{:02}.2024", (index % 12) + 1)),
        report_end: Some("30.06.2024".to_string()),
        table,
        positions,
        body_positions: cat_counts.into_iter().collect(),
        footer_bank: Some(base_cost * 10.0),
        footer_kasse: Some(base_cost * 2.0),
        footer_sonstiges: Some(base_cost * 0.5),
        locked: true,
        hide_columns_qv: true,
        ..ReportConfig::default()
    }
}

// ============================================================================
// Messung
// ============================================================================

/// Führt eine Batch-Generierung durch und gibt die Dauer zurück.
fn run_single(count: usize, output_dir: &Path, offset: usize) -> Duration {
    let start = Instant::now();
    for i in 0..count {
        let config = build_config(offset + i);
        let path = output_dir.join(format!("r_{:07}.xlsx", offset + i));
        config.write_to(&path).ok();
    }
    start.elapsed()
}

fn run_multi(count: usize, output_dir: &Path, num_threads: usize, offset: usize) -> Duration {
    let start = Instant::now();
    let output_dir = Arc::new(output_dir.to_path_buf());
    let chunk_size = count.div_ceil(num_threads);

    let handles: Vec<_> = (0..num_threads)
        .map(|tid| {
            let dir = Arc::clone(&output_dir);
            let start_i = offset + tid * chunk_size;
            let end_i = (offset + count).min(start_i + chunk_size);
            thread::spawn(move || {
                for i in start_i..end_i {
                    let config = build_config(i);
                    let path = dir.join(format!("r_{:07}.xlsx", i));
                    config.write_to(&path).ok();
                }
            })
        })
        .collect();

    for h in handles {
        h.join().unwrap();
    }
    start.elapsed()
}

fn run_multi_protected(
    count: usize,
    output_dir: &Path,
    num_threads: usize,
    hash: &Arc<kmw_fb_rust::PrecomputedHash>,
    offset: usize,
) -> Duration {
    let start = Instant::now();
    let output_dir = Arc::new(output_dir.to_path_buf());
    let chunk_size = count.div_ceil(num_threads);

    let handles: Vec<_> = (0..num_threads)
        .map(|tid| {
            let dir = Arc::clone(&output_dir);
            let hash = Arc::clone(hash);
            let start_i = offset + tid * chunk_size;
            let end_i = (offset + count).min(start_i + chunk_size);
            thread::spawn(move || {
                for i in start_i..end_i {
                    let config = build_config(i);
                    let path = dir.join(format!("r_{:07}.xlsx", i));
                    config.write_to_precomputed(&path, &hash).ok();
                }
            })
        })
        .collect();

    for h in handles {
        h.join().unwrap();
    }
    start.elapsed()
}

/// Misst einen Benchmark mit Warmup und N Samples.
/// Gibt Throughput-Samples (Dateien/Sekunde) zurück.
fn measure<F>(count: usize, warmup: usize, samples: usize, dir: &Path, run_fn: F) -> BenchStats
where
    F: Fn(&Path, usize) -> Duration,
{
    // Warmup (verwerfen)
    for w in 0..warmup {
        let offset = w * count;
        run_fn(dir, offset);
        // Dateien löschen um Disk-Füllstand konstant zu halten
        cleanup_dir(dir);
    }

    // Samples messen
    let mut throughput_samples = Vec::with_capacity(samples);
    for s in 0..samples {
        let offset = (warmup + s) * count;
        let duration = run_fn(dir, offset);
        let throughput = count as f64 / duration.as_secs_f64();
        throughput_samples.push(throughput);
        cleanup_dir(dir);
    }

    BenchStats::new(throughput_samples)
}

fn cleanup_dir(dir: &Path) {
    if dir.exists() {
        fs::remove_dir_all(dir).ok();
    }
    fs::create_dir_all(dir).ok();
}

// ============================================================================
// Alle Benchmarks ausführen
// ============================================================================

fn run_all_benchmarks(base_dir: &Path) -> Vec<BenchResult> {
    let mut results = Vec::new();
    let cpu_count = thread::available_parallelism().map(|p| p.get()).unwrap_or(4);

    for &(file_count, warmup, samples, multi_only) in BENCH_CONFIGS {
        let dir = base_dir.join(format!("bench_{}", file_count));
        fs::create_dir_all(&dir).unwrap();

        // Single-threaded (außer bei multi_only)
        if !multi_only {
            println!(
                "\n  [{} files] Single-threaded  (warmup={}, samples={})",
                fmt_count(file_count),
                warmup,
                samples
            );
            print!("    ");
            let stats = measure(file_count, warmup, samples, &dir, |d, off| {
                print!(".");
                std::io::Write::flush(&mut std::io::stdout()).ok();
                run_single(file_count, d, off)
            });
            println!(
                " {:.0}/sec  (CV {:.1}% — {})",
                stats.mean(),
                stats.cv_percent(),
                stats.quality()
            );
            results.push(BenchResult {
                label: format!("{} files", fmt_count(file_count)),
                file_count,
                threads: 1,
                stats,
            });
        }

        // Multi-threaded
        let active_threads: Vec<usize> = THREAD_COUNTS
            .iter()
            .copied()
            .filter(|&t| t <= cpu_count * 2)
            .collect();

        for num_threads in active_threads {
            println!(
                "\n  [{} files] {}-threaded  (warmup={}, samples={})",
                fmt_count(file_count),
                num_threads,
                warmup,
                samples
            );
            print!("    ");
            let stats = measure(file_count, warmup, samples, &dir, |d, off| {
                print!(".");
                std::io::Write::flush(&mut std::io::stdout()).ok();
                run_multi(file_count, d, num_threads, off)
            });
            println!(
                " {:.0}/sec  (CV {:.1}% — {})",
                stats.mean(),
                stats.cv_percent(),
                stats.quality()
            );
            results.push(BenchResult {
                label: format!("{} files", fmt_count(file_count)),
                file_count,
                threads: num_threads,
                stats,
            });
        }

        fs::remove_dir_all(&dir).ok();
    }

    // Protected benchmarks
    let prot_file_count = 1_000;
    let prot_warmup = 2;
    let prot_samples = 5;

    for &spin_count in &[1_000u32, 100_000u32] {
        let label = if spin_count == 1_000 { "fast" } else { "standard" };
        let dir = base_dir.join(format!("bench_prot_{}", spin_count));
        fs::create_dir_all(&dir).unwrap();
        let hash = Arc::new(precompute_hash_with_spin_count("benchmark_password", spin_count));

        println!(
            "\n  [{} files, {} threads, protection {}] (warmup={}, samples={})",
            fmt_count(prot_file_count),
            PROT_THREADS,
            label,
            prot_warmup,
            prot_samples
        );
        print!("    ");

        let hash_ref = Arc::clone(&hash);
        let stats = measure(
            prot_file_count,
            prot_warmup,
            prot_samples,
            &dir,
            move |d, off| {
                print!(".");
                std::io::Write::flush(&mut std::io::stdout()).ok();
                run_multi_protected(prot_file_count, d, PROT_THREADS, &hash_ref, off)
            },
        );
        println!(
            " {:.0}/sec  (CV {:.1}% — {})",
            stats.mean(),
            stats.cv_percent(),
            stats.quality()
        );
        results.push(BenchResult {
            label: format!("1,000 files + protection (spin={})", spin_count),
            file_count: prot_file_count,
            threads: PROT_THREADS,
            stats,
        });

        fs::remove_dir_all(&dir).ok();
    }

    results
}

// ============================================================================
// Ausgabe
// ============================================================================

fn print_summary(results: &[BenchResult], hw: &HardwareInfo) {
    println!("\n\n{}", "=".repeat(70));
    println!("  ERGEBNISSE");
    println!("  {}", hw.summary());
    println!("{}", "=".repeat(70));

    // Standard-Benchmarks
    let standard: Vec<_> = results
        .iter()
        .filter(|r| !r.label.contains("protection"))
        .collect();

    if !standard.is_empty() {
        println!(
            "\n  {:<22} {:>7} {:>10} {:>8} {:>7} {:>20} {:>8}",
            "Configuration", "Threads", "Median", "Std Dev", "CV", "Range (files/sec)", "Throughput"
        );
        println!("  {}", "-".repeat(90));
        for r in &standard {
            let dur = Duration::from_secs_f64(r.file_count as f64 / r.stats.median());
            let std_dur = Duration::from_secs_f64(
                r.stats.std_dev() * r.file_count as f64 / r.stats.mean().powi(2),
            );
            println!(
                "  {:<22} {:>7} {:>10} {:>8} {:>6.1}% {:>9.0}–{:<9.0} {:>8.0}/sec",
                r.label,
                r.threads,
                fmt_duration(dur),
                format!("±{}", fmt_duration(std_dur)),
                r.stats.cv_percent(),
                r.stats.min(),
                r.stats.max(),
                r.throughput_mean(),
            );
        }
    }

    // Protected benchmarks
    let protected: Vec<_> = results
        .iter()
        .filter(|r| r.label.contains("protection"))
        .collect();

    if !protected.is_empty() {
        println!("\n  Protection benchmarks (precomputed hash):");
        println!(
            "  {:<38} {:>7} {:>10} {:>7} {:>20} {:>8}",
            "Configuration", "Threads", "Median", "CV", "Range (files/sec)", "Throughput"
        );
        println!("  {}", "-".repeat(90));
        for r in &protected {
            let dur = Duration::from_secs_f64(r.file_count as f64 / r.stats.median());
            println!(
                "  {:<38} {:>7} {:>10} {:>6.1}% {:>9.0}–{:<9.0} {:>8.0}/sec",
                r.label,
                r.threads,
                fmt_duration(dur),
                r.stats.cv_percent(),
                r.stats.min(),
                r.stats.max(),
                r.throughput_mean(),
            );
        }
    }

    println!("\n  CV legend: excellent (<5%)  good (5–10%)  fair (10–20%)  noisy (>20%)");
    println!("{}", "=".repeat(70));
}

// ============================================================================
// README-Update
// ============================================================================

const PERF_START: &str = "<!-- PERF_START -->";
const PERF_END: &str = "<!-- PERF_END -->";
const README_PATH: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/README.md");

fn generate_perf_markdown(results: &[BenchResult], hw: &HardwareInfo) -> String {
    let mut out = String::new();

    out.push_str(&format!("**Environment:** {}\n\n", hw.summary()));

    // Haupt-Tabelle
    let standard: Vec<_> = results
        .iter()
        .filter(|r| !r.label.contains("protection"))
        .collect();

    if !standard.is_empty() {
        out.push_str("| Files | Threads | Mean | Std Dev | CV | Throughput |\n");
        out.push_str("|-------|---------|------|---------|----|------------|\n");

        for r in &standard {
            let dur = r.stats.mean_duration(r.file_count);
            let std_dur = Duration::from_secs_f64(
                r.stats.std_dev() * r.file_count as f64 / r.stats.mean().powi(2),
            );
            out.push_str(&format!(
                "| {:>6} | {:>7} | {:>8} | {:>8} | {:>4.1}% | {:>8.0}/sec |\n",
                fmt_count(r.file_count),
                r.threads,
                fmt_duration(dur),
                format!("±{}", fmt_duration(std_dur)),
                r.stats.cv_percent(),
                r.throughput_mean(),
            ));
        }
        out.push('\n');
    }

    // Protected-Tabelle
    let protected: Vec<_> = results
        .iter()
        .filter(|r| r.label.contains("protection"))
        .collect();

    if !protected.is_empty() {
        out.push_str(&format!(
            "**With workbook protection (precomputed hash, 1,000 files, {} threads):**\n\n",
            PROT_THREADS
        ));
        out.push_str("| Spin Count | Mean | CV | Throughput |\n");
        out.push_str("|------------|------|----|------------|\n");
        for r in &protected {
            let spin = if r.label.contains("1000") { "1,000" } else { "100,000" };
            let dur = r.stats.mean_duration(r.file_count);
            out.push_str(&format!(
                "| {:>10} | {:>8} | {:>4.1}% | {:>8.0}/sec |\n",
                spin,
                fmt_duration(dur),
                r.stats.cv_percent(),
                r.throughput_mean(),
            ));
        }
        out.push('\n');
    }

    // Datum
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| {
            let secs = d.as_secs();
            let days = secs / 86400;
            let year = 1970 + days / 365;
            let day_of_year = days % 365;
            let month = (day_of_year / 30) + 1;
            let day = (day_of_year % 30) + 1;
            format!("{:04}-{:02}-{:02}", year, month.min(12), day.min(31))
        })
        .unwrap_or_else(|_| "unknown".to_string());

    out.push_str(&format!("*Last updated: {}*\n", now));
    out
}

fn update_readme(results: &[BenchResult], hw: &HardwareInfo) {
    let readme_path = PathBuf::from(README_PATH);

    let content = match fs::read_to_string(&readme_path) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("\n  Warning: README.md not found or unreadable: {}", e);
            return;
        }
    };

    let start_pos = match content.find(PERF_START) {
        Some(p) => p,
        None => {
            eprintln!("\n  Warning: {} marker not found in README.md", PERF_START);
            return;
        }
    };
    let end_pos = match content.find(PERF_END) {
        Some(p) => p,
        None => {
            eprintln!("\n  Warning: {} marker not found in README.md", PERF_END);
            return;
        }
    };

    let new_content = generate_perf_markdown(results, hw);
    let updated = format!(
        "{}{}\n{}{}\n{}",
        &content[..start_pos],
        PERF_START,
        new_content,
        PERF_END,
        &content[end_pos + PERF_END.len()..]
    );

    match fs::write(&readme_path, updated) {
        Ok(_) => println!("\n  README.md updated."),
        Err(e) => eprintln!("\n  Warning: could not write README.md: {}", e),
    }
}

// ============================================================================
// Hilfsfunktionen
// ============================================================================

fn fmt_count(n: usize) -> String {
    if n >= 1_000 {
        format!("{},{:03}", n / 1_000, n % 1_000)
    } else {
        n.to_string()
    }
}

fn fmt_duration(d: Duration) -> String {
    let ms = d.as_secs_f64() * 1000.0;
    if ms >= 1_000.0 {
        format!("{:.2}s", ms / 1000.0)
    } else {
        format!("{:.0}ms", ms)
    }
}

// ============================================================================
// main
// ============================================================================

fn main() {
    println!("\n{}", "#".repeat(70));
    println!("  FINANZBERICHT BENCHMARK");
    println!("  Warmup + Statistical Sampling");
    println!("{}", "#".repeat(70));

    // Hardware erkennen
    print!("\n  Detecting hardware... ");
    std::io::Write::flush(&mut std::io::stdout()).ok();
    let hw = HardwareInfo::collect();
    println!("{}", hw.summary());

    let cpu_count = thread::available_parallelism().map(|p| p.get()).unwrap_or(4);
    println!("  Max threads tested: min({}, {})", cpu_count * 2, THREAD_COUNTS.last().unwrap());
    println!("\n  Note: Run with --release for representative results.");
    println!("  Methodology: Throughput (files/sec) measured per sample.");
    println!("               CV = coefficient of variation (lower = more stable).\n");

    let temp_dir = std::env::temp_dir().join("finanzbericht_benchmark");
    fs::create_dir_all(&temp_dir).unwrap();

    // Benchmarks ausführen
    let results = run_all_benchmarks(&temp_dir);

    // Zusammenfassung ausgeben
    print_summary(&results, &hw);

    // README aktualisieren
    update_readme(&results, &hw);

    // Temp-Verzeichnis aufräumen
    fs::remove_dir_all(&temp_dir).ok();

    println!("\n{}", "#".repeat(70));
    println!("  BENCHMARK ABGESCHLOSSEN");
    println!("{}\n", "#".repeat(70));
}
