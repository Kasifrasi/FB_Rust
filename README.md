# KMW-FB_Rust - Financial Report Generator

A high-performance **Excel financial report generator** written in Rust. Generates .xlsx files with dynamic cost positions, formulas, multi-language support, and professional formatting.

## ✨ Features

- 📊 **Excel Report Generation** - XLSX format with formulas and formatting
- 📈 **Dynamic Positions** - Unlimited cost positions per budget category
- 🌍 **Multi-Language** - 5 languages: German, English, French, Spanish, Portuguese
- ⚡ **High Performance** - ~260 files/second (single-threaded), ~1400 files/second (multi-threaded)
- 🔒 **Professional** - Sheet protection, validations, formatting
- 🧪 **Well-Tested** - 92 unit tests, integration tests, benchmarks

## 🚀 Quick Start

### Prerequisites
- Rust 1.70+ or NixOS with devenv
- ~1.2 GB disk space

### Build & Run

```bash
# Build release version
cargo build --release

# Run tests
cargo test --lib

# Generate example reports (5 languages)
cargo run --example test_multilang --release
# Output: tests/output/*.xlsx

# Performance benchmark
cargo run --example benchmark --release
```

### Output Location
All generated files write to: **`tests/output/`**
- Test artifacts: 10 files (217 KB)
- Example reports: 5 files (105 KB)

## 📚 Documentation

| Document | Purpose |
|----------|---------|
| **PROJECT_STRUCTURE.md** | Detailed architecture, module overview, code flow |
| **TESTFILES_CONSOLIDATION.md** | Test consolidation report, benchmark results |
| **This README** | Quick start guide |

## 🏗️ Architecture Overview

```
ReportValues (API Values)
    ↓
CellRegistry (Formulas + Topological Sort)
    ↓
BodyLayout (Position Calculations)
    ↓
Section Writers (header, table, panel, prebody)
    ↓
write_cells_from_registry (Formulas + Formats)
    ↓
Excel File (.xlsx)
```

## 📊 Performance

### Benchmark Results (12-core CPU, Release Build)

**100 Files:**
- Single-threaded: 263 files/sec
- 8 threads: 1,464 files/sec (5.55x speedup)
- 16 threads: 1,535 files/sec (5.82x speedup)

**1000 Files:**
- Single-threaded: 259 files/sec
- 8 threads: 1,231 files/sec (4.76x speedup)
- 16 threads: 1,401 files/sec (5.42x speedup)

**Optimal:** 8 threads for 12-core CPU

## 🧪 Testing

```bash
# All tests (92 total)
cargo test --lib

# Specific test suite
cargo test v2::tests::integrated_api
cargo test v2::tests::legacy
cargo test v2::tests_right_panel

# Individual test
cargo test test_realistic_medium_project

# With output
cargo test --lib -- --nocapture
```

### Test Coverage
- **body_config** - 5 tests (configuration validation)
- **body_layout** - 9 tests (layout calculations)
- **integrated_api** - 5 tests (realistic scenarios)
- **legacy** - 3 tests (header generation, formulas)
- **position_api** - 13 tests (position API)
- **tests_right_panel** - 2 tests (right panel)
- **Other** - 48 tests (formats, registry, utilities)

## 📁 Project Structure

```
src/
├── lib.rs                    [Entry point]
└── v2/                       [Current version]
    ├── common.rs             [Utilities]
    ├── lang/                 [Language support]
    ├── report/               [Main engine]
    │   ├── api.rs            [Public API]
    │   ├── registry.rs       [Formula management]
    │   ├── writer.rs         [Main write function]
    │   ├── body/             [Dynamic positions]
    │   └── sections/         [Static sections]
    └── tests/                [Unit tests (35 tests)]

tests/output/                [Test artifacts (340 KB)]

examples/
├── benchmark.rs             [Performance benchmarks]
└── test_multilang.rs        [Multilingual examples]
```

See **PROJECT_STRUCTURE.md** for detailed architecture.

## 💡 Usage Example

```rust
use kmw_fb_rust::v2::report::{
    write_report_v2_with_body, ApiKey, BodyConfig, ReportValues, ReportStyles
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let ws = workbook.add_worksheet();
    
    let styles = ReportStyles::new();
    let mut values = ReportValues::new()
        .with_language("deutsch")
        .with_currency("EUR")
        .with_project_number("KMW-2024-001")
        .with_project_title("Sample Project");
    
    // Add cost positions
    values.set_position_row(1, 1, "Position 1.1", 10000.0, 5000.0, 5000.0, "");
    values.set_position_row(1, 2, "Position 1.2", 8000.0, 4000.0, 4000.0, "");
    
    let body_config = BodyConfig::new()
        .with_positions(1, 2)  // Category 1: 2 positions
        .with_positions(2, 3)  // Category 2: 3 positions
        .with_positions(6, 0)  // Category 6: header-input mode
        .with_positions(7, 0)  // Category 7: header-input mode
        .with_positions(8, 0); // Category 8: header-input mode
    
    write_report_v2_with_body(ws, &styles, "_de", &values, &body_config)?;
    workbook.save("report.xlsx")?;
    
    Ok(())
}
```

## 🔧 Development

### Setup with DevEnv (NixOS)
```bash
devenv up              # Setup environment
devenv shell           # Enter shell
cargo build --release
```

### Setup without DevEnv
```bash
rustup update          # Update Rust
cargo build --release
```

### Watch Mode
```bash
cargo watch -x build -x test
```

### Clean Build
```bash
cargo clean
rm -rf tests/output/*.xlsx
cargo build --release
```

## 📈 Benchmarking

```bash
cargo run --example benchmark --release
```

Generates 100 and 1000 test reports, measures:
- Single-threaded performance
- Multi-threaded performance (2, 4, 8, 16 threads)
- Speedup efficiency
- Files per second

Results saved to stdout.

## 🎯 Key Files

| File | Purpose |
|------|---------|
| `src/v2/report/api.rs` | Public API & ApiKey enum |
| `src/v2/report/writer.rs` | Main write_report_v2() function |
| `src/v2/report/registry.rs` | Central formula management |
| `src/v2/report/definitions.rs` | Formula registrations |
| `examples/test_multilang.rs` | Working example |
| `examples/benchmark.rs` | Performance test |

## 📝 Recent Updates

**2026-01-20 - Major Cleanup**
- ✅ Deleted V1 legacy code (88 KB)
- ✅ Removed old output directories (126 KB)
- ✅ Cleaned development cache (1.3 MB)
- ✅ Consolidated all test outputs to `tests/output/`
- ✅ Created comprehensive PROJECT_STRUCTURE.md
- ✅ Total cleanup: ~1.6 MB

## 🚨 Troubleshooting

### Tests fail with "file not found"
- Make sure `tests/output/` directory exists
- Run: `mkdir -p tests/output`

### Build fails
- Update Rust: `rustup update`
- Clean build: `cargo clean && cargo build --release`

### Performance is slow
- Use release build: `cargo build --release`
- Check CPU cores: Optimal performance at ~8 threads per 12 cores

## 📞 Support

1. **Check documentation:**
   - `PROJECT_STRUCTURE.md` - Architecture details
   - `TESTFILES_CONSOLIDATION.md` - Test info & benchmarks
   - `examples/` - Working code samples

2. **Run tests:**
   ```bash
   cargo test --lib -- --nocapture
   ```

3. **Check logs:**
   - Test output includes file paths and counts
   - Benchmark output shows thread scaling

## 📄 License

[Project License Info]

---

**Status:** ✅ Production Ready  
**Version:** 2.0 (V1 deprecated)  
**Last Updated:** 2026-01-20  
**Lines of Code:** ~12,000 (V2 only)
