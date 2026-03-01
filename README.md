# FB-Rust - Financial Report Generator

High-performance Excel financial report generator written in Rust.

## Features

- Excel generation (XLSX) with formulas and formatting
- Dynamic cost positions per budget category
- Multi-language support (German, English, French, Spanish, Portuguese)
- High performance (~1200 reports/second multi-threaded)
- Sheet and workbook protection with locked formulas
- Typed error handling (`ReportError`, `ProtectionError`)
- Optional serde support (`--features serde`) with validated deserialization

## Quick Start

```bash
cargo build --release
cargo test
cargo run --example test_all_fields --release
cargo run --example test_multilang --release

# Criterion-Benchmarks (Statistik + HTML-Report + README-Update)
# Run the benchmarks
cargo bench

# HTML Report location:
# target/criterion/report/index.html

# For maximum reproducibility: Tune the system beforehand
# sudo pyperf system tune && cargo bench && sudo pyperf system reset
```

## Usage

### Simple API (recommended)

`ReportConfig` is the main entry point — composed of four domain sub-structs:
`ReportHeader`, `ReportBody`, `ReportFooter`, and `ReportOptions`.

All structs support the Builder pattern via `derive_builder`:

```rust
use fb_rust::*;

let config = ReportConfigBuilder::default()
    .header(
        ReportHeaderBuilder::default()
            .language(Language::Deutsch)
            .currency(Currency::EUR)
            .project_number("PROJ-2025-001")
            .project_title("Bildungsprojekt Ostafrika")
            .project_start("01.01.2025")
            .project_end("31.12.2027")
            .report_start("01.01.2025")
            .report_end("30.06.2025")
            .build()?
    )
    .body(
        ReportBodyBuilder::default()
            .add_table_entry(
                TableEntryBuilder::default()
                    .index(0u8)
                    .approved_budget(80000.0)
                    .income_report(50000.0)
                    .income_total(50000.0)
                    .reason("1. Rate")
                    .build()?
            )
            .add_position(
                PositionEntryBuilder::default()
                    .category(1u8)
                    .position(1u16)
                    .description("Projektleitung")
                    .approved(18000.0)
                    .income_report(9000.0)
                    .income_total(9000.0)
                    .build()?
            )
            .body_positions(
                [(1u8, 5u16), (2, 3), (3, 4), (4, 3), (5, 2), (6, 0), (7, 0), (8, 0)]
                    .into_iter().collect::<std::collections::HashMap<_, _>>()
            )
            .build()?
    )
    .footer(
        ReportFooterBuilder::default()
            .bank(8500.0)
            .kasse(1200.0)
            .sonstiges(300.0)
            .build()?
    )
    .options(
        ReportOptionsBuilder::default()
            .sheet_password("blatt_geheim")
            .workbook_password("secret")
            .hide_columns_qv(true)
            .hide_language_sheet(true)
            .build()?
    )
    .build()?;

config.write_to("report.xlsx")?;
```

Or using struct literals directly:

```rust
use fb_rust::{Currency, Language, ReportConfig, ReportHeader, ReportOptions};

let config = ReportConfig {
    header: ReportHeader {
        language: Language::Deutsch,
        currency: Currency::EUR,
        project_number: Some("PROJ-2025-001".to_string()),
        ..ReportHeader::default()
    },
    options: ReportOptions {
        sheet_password: Some("geheim".to_string()),
        ..ReportOptions::default()
    },
    ..ReportConfig::default()
};
```

> See [`examples/test_all_fields.rs`](examples/test_all_fields.rs) for a complete reference with all fields populated.

## Project Structure

```
src/
├── lib.rs                  Public re-exports and crate documentation
├── config.rs               ReportConfig — main entry point (Tauri-ready)
├── error.rs                ReportError — typed top-level error enum
├── lang/                   Language data (TEXT_MATRIX, CURRENCIES) and sheet builder
├── workbook_protection.rs  ProtectionError + ZIP-level workbook structure lock (SHA-512)
└── report/
    ├── api/                ApiKey, ReportValues, Language, Currency, ReportDate, …
    ├── core/               Formula engine (CellRegistry, CellAddr, topological eval)
    ├── options.rs          SheetOptions, SheetProtection, validation rules
    ├── styles.rs           ReportStyles, FormatMatrix (internal)
    ├── body/               BodyConfig, BodyLayout, FooterLayout, formulas
    └── writer/             Excel writing (engine.rs, layout.rs, structure.rs)

tests/
└── serde_integration.rs    Serde JSON roundtrip tests (27 tests, feature-gated)

examples/
├── test_all_fields.rs         All fields populated (complete reference example)
├── test_multilang.rs          5 languages
├── test_multilang_grouped.rs  Row grouping / collapsed sections
├── test_workbook_protection.rs  3 protection levels
├── test_parallel_protection.rs  Parallel protection benchmark (Rayon)
├── profile.rs                 Phase-level profiling
└── verify_password.rs         Password hash verification
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
cargo test                                             # Unit tests (77 tests)
cargo test --features serde --test serde_integration   # Serde JSON tests (27 tests)
cargo test --features serde                            # All tests (104 tests)
cargo deny check                                       # License and security audit
```

## License

Licensed under either of

- [MIT License](LICENSE-MIT)
- [Apache License, Version 2.0](LICENSE-2.0)

at your option.
