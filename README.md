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
cargo run --example benchmark --release
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
    ├── format/             Styles, FormatMatrix, SheetProtection, validation rules
    ├── body/               BodyConfig, BodyLayout, FooterLayout, formulas
    └── writer/             Excel writing (main.rs, layout.rs, structure.rs)

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

| Files | Single-thread | 8 Threads | Speedup |
|-------|---------------|-----------|---------|
| 100   | 264/sec       | 1464/sec  | 5.55x   |
| 1000  | 259/sec       | 1231/sec  | 4.76x   |

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
