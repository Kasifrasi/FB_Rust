# KMW-FB Rust - Financial Report Generator

High-performance Excel financial report generator written in Rust.

## Features

- Excel generation (XLSX) with formulas and formatting
- Dynamic cost positions per budget category
- Multi-language support (German, English, French, Spanish, Portuguese)
- High performance (~1200 reports/second multi-threaded)
- Sheet and workbook protection with locked formulas

## Quick Start

```bash
cargo build --release
cargo test
cargo run --example test_all_fields --release
cargo run --example test_multilang --release

# Criterion-Benchmarks (Statistik + HTML-Report + README-Update)
cargo bench
# HTML-Report: target/criterion/report/index.html

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
├── test_parallel_protection.rs  Parallel protection benchmark (Rayon)
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

> **Methodology:** Measured with [criterion.rs](https://github.com/bheisler/criterion.rs) (bootstrap CI, outlier detection).
> CV = coefficient of variation (σ/μ) — lower means more stable results.
> Run `cargo bench` to reproduce on your hardware.
>
> **Benchmark config:** Each report uses all API features: 5 table rows, 2×18 panel entries,
> 50 cost positions (categories 1–5 with 10 each) + 3 header-input categories (6–8),
> sheet protection, workbook protection (SHA-512, spin count 1,000 per batch),
> hidden columns, hidden language sheet, row grouping, and all footer fields.

<!-- PERF_START -->
**Environment:** AMD Ryzen 5 Pro 7535U with Radeon Graphics · 6 cores (12 logical) · 30 GB RAM · Linux (NixOS 26.05)

| Files | Threads | Mean | Std Dev | CV | Throughput |
|------:|--------:|-----:|--------:|---:|-----------:|
|   100 |       1 |    377ms |      ±4ms |  1.0% |        265/sec |
|   100 |       8 |     87ms |      ±1ms |  1.4% |       1154/sec |
|   100 |      16 |     80ms |      ±2ms |  2.9% |       1246/sec |
| 1,000 |       1 |    3.88s |     ±31ms |  0.8% |        258/sec |
| 1,000 |       8 |    867ms |     ±38ms |  4.4% |       1153/sec |
| 1,000 |      16 |    838ms |     ±37ms |  4.4% |       1193/sec |

*Last updated: 2026-03-01*
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
