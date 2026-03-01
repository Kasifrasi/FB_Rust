//! Criterion-Benchmarks: Finanzbericht-Generierung
//!
//! Misst Throughput (Dateien/Sekunde) für Single- und Multi-threaded
//! Generierung mit statistischer Auswertung, Outlier-Erkennung und
//! HTML-Reports.
//!
//! Usage:
//!   cargo bench
//!   cargo bench -- --save-baseline main    # Baseline speichern
//!   cargo bench -- --baseline main         # Gegen Baseline vergleichen
//!
//! HTML-Reports: target/criterion/report/index.html
//!
//! Für maximale Reproduzierbarkeit vorher System tunen:
//!   sudo pyperf system tune
//!   cargo bench
//!   sudo pyperf system reset

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use kmw_fb_rust::{PositionEntry, ReportConfig, TableEntry};
use std::fs;
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant};

// ============================================================================
// Test-Config-Generierung
// ============================================================================

fn build_config(index: usize) -> ReportConfig {
    const LANGUAGES: [&str; 5] = ["deutsch", "english", "francais", "espanol", "portugues"];
    const CURRENCIES: [&str; 5] = ["EUR", "USD", "GBP", "CHF", "BRL"];

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
        language: LANGUAGES[index % 5].to_string(),
        currency: CURRENCIES[index % 5].to_string(),
        project_number: Some(format!("PROJ-{:05}", index)),
        project_title: Some(format!("Project {} ({})", index, LANGUAGES[index % 5])),
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

fn cleanup_dir(dir: &std::path::Path) {
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            fs::remove_file(entry.path()).ok();
        }
    }
}

// ============================================================================
// Benchmarks
// ============================================================================

/// Single-threaded Throughput: 100 und 1.000 Dateien
fn bench_single_threaded(c: &mut Criterion) {
    let dir = tempfile::TempDir::new().unwrap();
    let mut group = c.benchmark_group("single_threaded");
    group
        .warm_up_time(Duration::from_secs(5))
        .measurement_time(Duration::from_secs(30));

    for &count in &[100usize, 1_000] {
        group.throughput(Throughput::Elements(count as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(count),
            &count,
            |b, &count| {
                b.iter_custom(|iters| {
                    let mut total = Duration::ZERO;
                    for i in 0..iters as usize {
                        let start = Instant::now();
                        for j in 0..count {
                            let config = build_config(i * count + j);
                            let path = dir.path().join(format!("r_{}.xlsx", j));
                            config.write_to(&path).ok();
                        }
                        total += start.elapsed();
                        cleanup_dir(dir.path());
                    }
                    total
                });
            },
        );
    }
    group.finish();
}

/// Multi-threaded Throughput mit 8 Threads: 100 und 1.000 Dateien
fn bench_multi_threaded_8(c: &mut Criterion) {
    const NUM_THREADS: usize = 8;
    let dir = tempfile::TempDir::new().unwrap();
    let mut group = c.benchmark_group("multi_threaded_8");
    group
        .warm_up_time(Duration::from_secs(5))
        .measurement_time(Duration::from_secs(30));

    for &count in &[100usize, 1_000] {
        group.throughput(Throughput::Elements(count as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(count),
            &count,
            |b, &count| {
                b.iter_custom(|iters| {
                    let mut total = Duration::ZERO;
                    for i in 0..iters as usize {
                        let dir_arc = Arc::new(dir.path().to_path_buf());
                        let chunk = count.div_ceil(NUM_THREADS);

                        let start = Instant::now();
                        let handles: Vec<_> = (0..NUM_THREADS)
                            .map(|tid| {
                                let d = Arc::clone(&dir_arc);
                                let start_j = tid * chunk;
                                let end_j = count.min(start_j + chunk);
                                thread::spawn(move || {
                                    for j in start_j..end_j {
                                        let config = build_config(i * count + j);
                                        let path =
                                            d.join(format!("r_{}_{}.xlsx", tid, j));
                                        config.write_to(&path).ok();
                                    }
                                })
                            })
                            .collect();
                        for h in handles {
                            h.join().ok();
                        }
                        total += start.elapsed();
                        cleanup_dir(dir.path());
                    }
                    total
                });
            },
        );
    }
    group.finish();
}

/// Vergleich: Single vs. 2 / 4 / 8 / 16 Threads für 1.000 Dateien
fn bench_thread_scaling(c: &mut Criterion) {
    const COUNT: usize = 1_000;
    let dir = tempfile::TempDir::new().unwrap();
    let mut group = c.benchmark_group("thread_scaling");
    group
        .warm_up_time(Duration::from_secs(3))
        .measurement_time(Duration::from_secs(20))
        .throughput(Throughput::Elements(COUNT as u64));

    let max_threads = thread::available_parallelism().map(|p| p.get()).unwrap_or(8);

    for num_threads in [1usize, 2, 4, 8, 16].iter().filter(|&&t| t <= max_threads * 2) {
        let num_threads = *num_threads;
        group.bench_with_input(
            BenchmarkId::new("threads", num_threads),
            &num_threads,
            |b, &num_threads| {
                b.iter_custom(|iters| {
                    let mut total = Duration::ZERO;
                    for i in 0..iters as usize {
                        let dir_arc = Arc::new(dir.path().to_path_buf());
                        let chunk = COUNT.div_ceil(num_threads);

                        let start = Instant::now();
                        let handles: Vec<_> = (0..num_threads)
                            .map(|tid| {
                                let d = Arc::clone(&dir_arc);
                                let start_j = tid * chunk;
                                let end_j = COUNT.min(start_j + chunk);
                                thread::spawn(move || {
                                    for j in start_j..end_j {
                                        let config = build_config(i * COUNT + j);
                                        let path =
                                            d.join(format!("r_{}_{}.xlsx", tid, j));
                                        config.write_to(&path).ok();
                                    }
                                })
                            })
                            .collect();
                        for h in handles {
                            h.join().ok();
                        }
                        total += start.elapsed();
                        cleanup_dir(dir.path());
                    }
                    total
                });
            },
        );
    }
    group.finish();
}

criterion_group!(
    benches,
    bench_single_threaded,
    bench_multi_threaded_8,
    bench_thread_scaling,
);
criterion_main!(benches);
