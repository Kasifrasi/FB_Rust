# KMW-FB Rust - Financial Report Generator

High-performance Excel financial report generator written in Rust.

## Features

- Excel generation (XLSX) with formulas and formatting
- Dynamic cost positions per budget category
- Multi-language support (German, English, French, Spanish, Portuguese)
- High performance (~1400 reports/second multi-threaded)
- Sheet and workbook protection with locked formulas

## Quick Start

```bash
cargo build --release
cargo test
cargo run --example test_all_fields --release
cargo run --example test_multilang --release

# Criterion-Benchmarks (Statistik + HTML-Report)
cargo bench
# Report öffnen: target/criterion/report/index.html

# Throughput-Benchmark + README-Update
cargo run --example benchmark --release

# Für maximale Reproduzierbarkeit: System-Tuning vorher
# sudo pyperf system tune && cargo bench && sudo pyperf system reset
```

## Usage

### Simple API (recommended)

`ReportConfig` is the main entry point — it bundles all parameters in one serializable struct:

```rust
use kmw_fb_rust::{ReportConfig, PositionEntry};

let config = ReportConfig {
    language: "deutsch".to_string(),
    currency: "EUR".to_string(),
    project_number: Some("PROJ-2025-001".to_string()),
    project_title: Some("Bildungsprojekt Ostafrika".to_string()),
    locked: true,
    hide_columns_qv: true,
    positions: vec![
        PositionEntry {
            category: 1,
            position: 1,
            description: Some("Personalkosten".to_string()),
            approved: Some(18000.0),
            income_report: Some(9000.0),
            income_total: Some(9000.0),
            remark: None,
        },
    ],
    body_positions: [(1u8, 5u16), (2, 3), (6, 0), (7, 0), (8, 0)]
        .into_iter()
        .collect(),
    footer_bank: Some(8500.0),
    ..ReportConfig::default()
};

config.write_to("report.xlsx")?;
```

## Project Structure

```
src/
├── lib.rs                  Public re-exports and crate documentation
├── config.rs               ReportConfig — main entry point (Tauri-ready)
├── lang/                   Language data (TEXT_MATRIX, CURRENCIES) and sheet builder
├── workbook_protection.rs  ZIP-level workbook structure lock (SHA-512)
└── report/
    ├── api/                ApiKey, ReportValues, Language, Currency, ReportDate, …
    ├── core/               Formula engine (CellRegistry, CellAddr, topological eval)
    ├── options.rs          ReportOptions, SheetProtection, validation rules
    ├── styles.rs           ReportStyles, FormatMatrix (internal)
    ├── body/               BodyConfig, BodyLayout, FooterLayout, formulas
    └── writer/             Excel writing (engine.rs, layout.rs, structure.rs)

examples/
├── test_all_fields.rs         All fields populated (complete reference example)
├── test_multilang.rs          5 languages
├── test_multilang_grouped.rs  Row grouping / collapsed sections
├── test_workbook_protection.rs  3 protection levels
├── benchmark.rs               Performance benchmark (single- vs multi-threaded)
├── profile.rs                 Phase-level profiling
└── verify_password.rs         Password hash verification
```

## BodyConfig

Categories 1–8 can each be configured independently:

```rust
let config = BodyConfig::new()
    .with_positions(1, 10)  // Category 1: 10 position rows under header
    .with_positions(6, 0);  // Category 6: header-input mode (no sub-rows)
```

## Performance

> **Methodology:** Throughput (files/sec) measured across multiple samples after warmup.
> CV = coefficient of variation (σ/μ) — lower means more stable results.
> Run `cargo run --example benchmark --release` to reproduce on your hardware.

<!-- PERF_START -->
**Environment:** AMD Ryzen 5 Pro 7535U with Radeon Graphics · 6 cores (12 logical) · 30 GB RAM · Linux (NixOS 26.05)

| Files | Threads | Mean | Std Dev | CV | Throughput |
|-------|---------|------|---------|----|------------|
|    100 |       1 |    344ms |    ±17ms |  5.1% |      291/sec |
|    100 |       2 |    162ms |     ±3ms |  1.9% |      616/sec |
|    100 |       4 |     89ms |     ±6ms |  6.7% |     1129/sec |
|    100 |       8 |     70ms |     ±2ms |  2.7% |     1426/sec |
|    100 |      16 |     66ms |     ±2ms |  2.8% |     1523/sec |
|  1,000 |       1 |    3.41s |    ±73ms |  2.1% |      293/sec |
|  1,000 |       2 |    1.66s |    ±14ms |  0.8% |      604/sec |
|  1,000 |       4 |    961ms |    ±60ms |  6.3% |     1040/sec |
|  1,000 |       8 |    716ms |    ±14ms |  2.0% |     1397/sec |
|  1,000 |      16 |    678ms |    ±10ms |  1.5% |     1474/sec |
| 10,000 |       2 |   17.81s |   ±796ms |  4.5% |      561/sec |
| 10,000 |       4 |   10.31s |   ±117ms |  1.1% |      970/sec |
| 10,000 |       8 |    7.50s |   ±117ms |  1.6% |     1334/sec |
| 10,000 |      16 |    7.60s |   ±138ms |  1.8% |     1315/sec |

**With workbook protection (precomputed hash, 1,000 files, 8 threads):**

| Spin Count | Mean | CV | Throughput |
|------------|------|----|------------|
|      1,000 |    925ms |  2.8% |     1081/sec |
|      1,000 |    923ms |  0.9% |     1084/sec |

*Last updated: 2026-03-14*
<!-- PERF_END -->


## Testing

```bash
cargo test             # All tests (unit + integration)
cargo test --lib       # Unit tests only
cargo deny check       # License and security audit
```

## License

Licensed under either of

- [MIT License](LICENSE-MIT)
- [Apache License, Version 2.0](LICENSE-2.0)

at your option.
