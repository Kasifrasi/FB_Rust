# KMW-FB Rust - Financial Report Generator

High-performance Excel financial report generator written in Rust.

## Features

- Excel generation (XLSX) with formulas and formatting
- Dynamic cost positions per budget category
- Multi-language support (German, English, French, Spanish, Portuguese)
- High performance (~1400 reports/second multi-threaded)
- Sheet protection with locked formulas

## Quick Start

```bash
cargo build --release
cargo test --lib
cargo run --example test_multilang --release
cargo run --example benchmark --release
```

## Usage

```rust
use kmw_fb_rust::{
    write_report_with_body, ApiKey, BodyConfig, ReportValues, ReportStyles
};

let mut workbook = rust_xlsxwriter::Workbook::new();
let ws = workbook.add_worksheet();
let styles = ReportStyles::new();

let mut values = ReportValues::new()
    .with_language("deutsch")
    .with_currency("EUR")
    .with_project_number("PROJECT-001")
    .with_project_title("Sample Project");

values.set_position_row(1, 1, "Personnel", 10000.0, 5000.0, 5000.0, "");

let body_config = BodyConfig::new()
    .with_positions(1, 1)
    .with_positions(6, 0)
    .with_positions(7, 0)
    .with_positions(8, 0);

write_report_with_body(ws, &styles, "_de", &values, &body_config)?;
workbook.save("report.xlsx")?;
```

## Project Structure

```
src/
├── lib.rs              Entry point
├── common.rs           Shared utilities
├── lang/               Language support
├── report/
│   ├── api.rs          Public API (ApiKey, ReportValues)
│   ├── registry.rs     Formula management
│   ├── writer.rs       Main write functions
│   ├── body/           Dynamic position sections
│   └── sections/       Static sections
└── tests/              Unit tests

examples/
├── benchmark.rs        Performance benchmarking
└── test_multilang.rs   Multi-language example
```

## API

### ReportValues

```rust
values.set(ApiKey::Language, "deutsch");
values.set(ApiKey::Currency, "EUR");
values.set(ApiKey::ApprovedBudget(0), 50000.0);
values.set_position_row(category, position, description, approved, reported, total, remark);
values.set_header_input(category, approved, reported, total, remark);
values.set_footer_bank(10000.0);
```

### BodyConfig

```rust
let config = BodyConfig::new()
    .with_positions(1, 10)  // Category 1: 10 positions
    .with_positions(6, 0);  // Category 6: header-input mode
```

## Performance

| Files | Single-thread | 8 Threads | Speedup |
|-------|---------------|-----------|---------|
| 100   | 264/sec       | 1464/sec  | 5.55x   |
| 1000  | 259/sec       | 1231/sec  | 4.76x   |

## Testing

```bash
cargo test --lib                        # All tests
cargo test tests::integrated_api        # Specific module
```

## License

[License information]
