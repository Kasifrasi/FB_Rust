# KMW-FB Rust - Financial Report Generator

High-performance Excel financial report generator written in Rust. Creates `.xlsx` files with dynamic cost positions, formulas, formatting, and multi-language support.

## Features

- **Excel Generation** - Native XLSX output with formulas and formatting
- **Dynamic Positions** - Configurable cost positions per budget category
- **Multi-Language** - German, English, French, Spanish, Portuguese
- **High Performance** - ~1400 reports/second (multi-threaded)
- **Sheet Protection** - Locked formulas, unlocked input cells

## Quick Start

```bash
# Build
cargo build --release

# Run tests
cargo test --lib

# Generate example reports
cargo run --example test_multilang --release

# Performance benchmark
cargo run --example benchmark --release
```

## Usage

```rust
use kmw_fb_rust::v2::report::{
    write_report_v2_with_body, ApiKey, BodyConfig, ReportValues, ReportStyles
};

let mut workbook = rust_xlsxwriter::Workbook::new();
let ws = workbook.add_worksheet();
let styles = ReportStyles::new();

// Configure report values
let mut values = ReportValues::new()
    .with_language("deutsch")
    .with_currency("EUR")
    .with_project_number("PROJECT-001")
    .with_project_title("Sample Project");

// Add cost positions
values.set_position_row(1, 1, "Personnel", 10000.0, 5000.0, 5000.0, "");
values.set_position_row(1, 2, "Equipment", 8000.0, 4000.0, 4000.0, "");

// Configure body layout
let body_config = BodyConfig::new()
    .with_positions(1, 2)   // Category 1: 2 positions
    .with_positions(2, 3)   // Category 2: 3 positions
    .with_positions(6, 0)   // Categories 6-8: header-input mode
    .with_positions(7, 0)
    .with_positions(8, 0);

// Generate report
write_report_v2_with_body(ws, &styles, "_de", &values, &body_config)?;
workbook.save("report.xlsx")?;
```

## Project Structure

```
src/
├── lib.rs                    Entry point
└── v2/
    ├── common.rs             Shared utilities
    ├── lang/                 Language support (VLOOKUP tables)
    ├── report/
    │   ├── api.rs            Public API (ApiKey, ReportValues)
    │   ├── registry.rs       Formula management with topological sort
    │   ├── writer.rs         Main write_report_v2() function
    │   ├── body/             Dynamic position sections
    │   └── sections/         Static sections (header, table, panel)
    └── tests/                Unit tests

examples/
├── benchmark.rs              Performance benchmarking
└── test_multilang.rs         Multi-language example

tests/output/                 Generated test artifacts
```

## Architecture

```
ReportValues + BodyConfig
        ↓
CellRegistry (formulas + topological sort)
        ↓
Section Writers (header, table, panel, prebody)
        ↓
Body Writer (dynamic positions)
        ↓
write_cells_from_registry (apply formulas + formats)
        ↓
Excel File (.xlsx)
```

## API Reference

### ReportValues

Container for all report data:

```rust
// Static values
values.set(ApiKey::Language, "deutsch");
values.set(ApiKey::Currency, "EUR");
values.set(ApiKey::ProjectNumber, "PROJECT-001");

// Income table (rows 0-4)
values.set(ApiKey::ApprovedBudget(0), 50000.0);
values.set(ApiKey::IncomeReportPeriod(0), 25000.0);

// Position data
values.set_position_row(category, position, description, approved, reported, total, remark);

// Header-input mode (categories 6-8 with position=0)
values.set_header_input(category, approved, reported, total, remark);

// Footer values
values.set_footer_bank(10000.0);
values.set_footer_kasse(2000.0);
values.set_footer_sonstiges(500.0);
```

### BodyConfig

Configure positions per budget category:

```rust
let config = BodyConfig::new()
    .with_positions(1, 10)  // Category 1: 10 positions
    .with_positions(2, 15)  // Category 2: 15 positions
    .with_positions(6, 0)   // Category 6: header-input mode (single row)
    .with_positions(7, 0)
    .with_positions(8, 0);
```

- Categories 1-5: Budget categories with multiple positions
- Categories 6-8: Special categories (position=0 for header-input mode)

## Performance

Benchmark results (12-core CPU, release build):

| Files | Single-threaded | 8 Threads | Speedup |
|-------|-----------------|-----------|---------|
| 100   | 264/sec         | 1464/sec  | 5.55x   |
| 1000  | 259/sec         | 1231/sec  | 4.76x   |

## Testing

```bash
# All tests (92 total)
cargo test --lib

# Specific module
cargo test v2::tests::integrated_api
cargo test v2::tests::body_layout
```

Test output files are written to `tests/output/`.

## Dependencies

- [rust_xlsxwriter](https://crates.io/crates/rust_xlsxwriter) - Excel generation
- [serde](https://crates.io/crates/serde) - Serialization
- [lazy_static](https://crates.io/crates/lazy_static) - Static initialization

## License

[License information]
