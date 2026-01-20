# KMW-FB_Rust - Projektstruktur

**Letzte Aktualisierung:** 2026-01-20  
**Version:** 2.0 (V1 entfernt, bereinigt und reorganisiert)

---

## 📋 Übersicht

KMW-FB_Rust ist ein **Finanzberichts-Generator** in Rust, der Excel-Dateien mit dynamischen Kostenpositionen, Formeln und Formatierung erstellt.

### Kernmerkmale
- ✅ Excel-Report-Generierung (XLSX)
- ✅ Dynamische Kostenpositionen (beliebige Anzahl pro Kategorie)
- ✅ Mehrsprachig (5 Sprachen: Deutsch, Englisch, Französisch, Spanisch, Portugiesisch)
- ✅ Formeln und Validierungen
- ✅ Formatierung und Blattschutz
- ✅ Performance: ~260 Dateien/Sekunde (Single-threaded), ~1400 Dateien/Sekunde (Multi-threaded)

---

## 🗂️ Verzeichnisstruktur

```
/home/ardit/repos/Arbeit/KMW-FB_Rust/src_rust/
│
├── 📁 src/                              [Quellcode]
│   ├── lib.rs                           [Library entry point]
│   └── 📁 v2/                           [Aktuelle Produktionsversion (V1 gelöscht)]
│       │
│       ├── 📄 Core Module
│       ├── mod.rs                       [V2 module entry point with legacy compat]
│       ├── common.rs                    [Gemeinsame Utilities]
│       │
│       ├── 📁 lang/                     [Sprachenunterstützung]
│       │   ├── mod.rs
│       │   ├── builder.rs               [Sprachsheet-Erstellung]
│       │   └── data.rs                  [Sprachdaten (VLOOKUP-Tabellen)]
│       │
│       ├── 📁 report/                   [Hauptreport-Engine]
│       │   ├── 🔵 API & Configuration
│       │   ├── mod.rs                   [Report module entry]
│       │   ├── api.rs                   [Public API (ApiKey enum, ReportValues)]
│       │   ├── dynamic.rs               [Dynamische Formelauswertung]
│       │   ├── formats.rs               [Report-Formatierung (ReportStyles)]
│       │   ├── values.rs                [Value Definitions & Storage]
│       │   ├── layout.rs                [Layout & Sheet-Setup]
│       │   │
│       │   ├── 🔵 Registry System
│       │   ├── registry.rs              [CellRegistry: Zentrale Formelverwaltung]
│       │   ├── formulas.rs              [Formel-Definitionen]
│       │   ├── cells.rs                 [Cell-Operationen]
│       │   ├── definitions.rs           [Alle Formel-Registrierungen]
│       │   │
│       │   ├── 🔵 Writer System
│       │   ├── writer.rs                [Main write_report_v2() function]
│       │   │
│       │   ├── 📁 body/                 [Dynamische Body-Sektion]
│       │   │   ├── mod.rs
│       │   │   ├── config.rs            [BodyConfig: Positionen pro Kategorie]
│       │   │   ├── layout.rs            [BodyLayout: Row-Berechnung]
│       │   │   ├── formulas.rs          [Body-Formeln registrieren]
│       │   │   ├── writer.rs            [Body-Daten schreiben]
│       │   │   └── footer.rs            [Footer-Sektionen]
│       │   │
│       │   └── 📁 sections/             [Statische Sektionen]
│       │       ├── mod.rs
│       │       ├── header.rs            [Header-Sektion (Projektinfo)]
│       │       ├── table.rs             [Einnahmen-Tabelle]
│       │       ├── panel.rs             [Rechtes Panel (K/R, O/V, L-U Spalten)]
│       │       └── prebody.rs           [Prebody-Sektion]
│       │
│       ├── 📁 tests/                    [Unit-Tests (35 Tests)]
│       │   ├── mod.rs                   [Test module entry]
│       │   ├── body_config.rs           [5 Tests: Body-Konfiguration]
│       │   ├── body_layout.rs           [9 Tests: Layout-Berechnung]
│       │   ├── integrated_api.rs        [5 Tests: Integration realistisch]
│       │   ├── legacy.rs                [3 Tests: Header & Formeln]
│       │   └── position_api.rs          [13 Tests: Position API]
│       │
│       ├── tests_right_panel.rs         [2 Tests: Right Panel]
│       │
│       └── 📁 Excel-Template/
│           └── template.xlsx            [Excel-Template-Referenz]
│
├── 📁 tests/                            [Integrations-Tests (zentrale Output)]
│   └── 📁 output/                       [Generierte Test-Excel-Dateien (340 KB)]
│       ├── header_test.xlsx             (Legacy header test)
│       ├── header_test_v2.xlsx          (Registry-based header test)
│       ├── formula_eval_test.xlsx       (Formula evaluation test)
│       ├── right_panel_*.xlsx           (Right panel tests - 2 files)
│       ├── test_realistic_*.xlsx        (Integration tests - 5 files)
│       ├── D_Finanzbericht.xlsx         (German example output)
│       ├── E_Financial_report.xlsx      (English example output)
│       ├── F_Rapport_financier.xlsx     (French example output)
│       ├── S_Informe_financiero.xlsx    (Spanish example output)
│       └── P_Relatorio_financeiro.xlsx  (Portuguese example output)
│
├── 📁 examples/                         [Runnable Examples (24 KB)]
│   ├── benchmark.rs                     [Performance Benchmark Suite]
│   │   • 100 & 1000 Dateien generieren
│   │   • Single- vs Multi-threaded vergleich
│   │   • Speedup-Messungen
│   │   • Temp-Dateien in /tmp/ (auto-cleanup)
│   └── test_multilang.rs                [Mehrsprachige Report-Generierung]
│       • Generiert 5 Excel-Dateien (1x pro Sprache)
│       • Output: tests/output/
│
├── 📁 target/                           [Build Artifacts (1.8 GB, auto-generated)]
│   ├── debug/                           [Debug builds]
│   └── release/                         [Optimized release builds]
│
├── 📁 .git/                             [Git Repository (history)]
│
├── 📁 .direnv/                          [Direnv cache (auto)]
│
├── 📄 Cargo.toml                        [Project manifest]
├── 📄 Cargo.lock                        [Dependency lock]
├── 📄 devenv.nix                        [NixOS environment definition]
├── 📄 devenv.lock                       [Devenv lock file]
├── 📄 devenv.yaml                       [Devenv configuration]
├── 📄 .envrc                            [Direnv configuration]
├── 📄 .gitignore                        [Git ignore rules]
├── 📄 .devenv.flake.nix                 [Flake definition]
│
├── 📚 Dokumentation
├── 📄 PROJECT_STRUCTURE.md              [Diese Datei]
├── 📄 TESTFILES_CONSOLIDATION.md        [Test Konsolidierungsbericht]
│
└── 📁 .vscode/                          [VS Code settings]
```

---

## 🔧 Core Architecture

### Module Hierarchy

```
lib.rs (entry point)
  └── v2/mod.rs (current version with legacy compat)
       ├── common.rs (utilities)
       ├── lang/ (language support)
       │   ├── builder.rs → build_sheet()
       │   └── data.rs → language strings
       │
       └── report/ (main engine)
           ├── 📦 API Layer
           │   ├── api.rs → ApiKey, ReportValues
           │   └── layout.rs → setup_sheet()
           │
           ├── 📦 Core Engine
           │   ├── dynamic.rs → formula evaluation
           │   ├── formats.rs → ReportStyles
           │   ├── values.rs → Value storage
           │   ├── cells.rs → cell operations
           │   └── registry.rs → CellRegistry (topological sort)
           │
           ├── 📦 Writer System
           │   ├── writer.rs → main write_report_v2()
           │   │   └── calls section writers
           │   │       ├── sections/header.rs
           │   │       ├── sections/table.rs
           │   │       ├── sections/panel.rs
           │   │       ├── sections/prebody.rs
           │   │       └── body/writer.rs
           │   │
           │   └── Registry Evaluation
           │       └── write_cells_from_registry()
           │           (applies formulas + formats)
           │
           ├── 📦 Body System
           │   ├── body/config.rs → BodyConfig
           │   ├── body/layout.rs → BodyLayout
           │   ├── body/formulas.rs → register_body_formulas()
           │   └── body/footer.rs → footer formulas
           │
           ├── 📦 Definitions
           │   ├── definitions.rs → register_formula_cells()
           │   │   (all static formulas)
           │   └── called from build_registry()
           │
           └── 📦 Tests
               ├── tests/mod.rs
               ├── tests/*.rs
               └── tests_right_panel.rs
```

### Data Flow

```
1. Input: ReportValues + BodyConfig
                ↓
2. Build Registry with topological sort
   (definitions.rs: register_formula_cells)
                ↓
3. Setup Worksheet Layout
   (header.rs, table.rs, panel.rs, prebody.rs)
                ↓
4. Build Body Layout
   (body/layout.rs: BodyLayout::compute)
                ↓
5. Evaluate All Formulas
   (registry.evaluate_all() with topological order)
                ↓
6. Write Cells + Formulas + Formats
   (write_cells_from_registry())
                ↓
7. Output: Excel File (.xlsx)
```

---

## 📊 Code Statistics

| Komponente | Dateien | Zeilen | Größe | Status |
|-----------|---------|--------|-------|--------|
| V2 Report Engine | 34 | 9,680 | 1.1 MB | ✅ Active |
| Unit Tests | 6 | N/A | N/A | ✅ 92 tests |
| Examples | 2 | N/A | 24 KB | ✅ 2 suites |
| **Total Source** | **42** | **~12K** | **~1.2 MB** | |
| **V1 (Legacy)** | **6** | **2,215** | **88 KB** | ❌ **DELETED** |

---

## 🧪 Testing

### Test Framework
- **Tool:** Rust `#[test]` macros
- **Test Count:** 92 tests
- **Location:** src/v2/tests/ + src/v2/tests_right_panel.rs
- **Coverage:** API, Layout, Config, Formulas, Integration

### Test Suites

#### 1. Unit Tests (35 tests in src/v2/tests/)
```
body_config.rs       (5 tests)  - BodyConfig validation
body_layout.rs       (9 tests)  - BodyLayout row calculations
integrated_api.rs    (5 tests)  - Realistic integration tests
legacy.rs            (3 tests)  - Header generation & formula eval
position_api.rs     (13 tests)  - Position API validation
```

#### 2. Right Panel Tests (2 tests)
```
tests_right_panel.rs (2 tests)  - L/M/N/S/T/U column tests
```

#### 3. Integration Tests (5 examples)
```
test_multilang.rs    - 5 language variants
                       (German, English, French, Spanish, Portuguese)
```

#### 4. Performance Benchmarks (2 suites)
```
benchmark.rs         - 100 & 1000 file generation
                       Single vs Multi-threaded comparison
                       Speedup measurements
```

### Running Tests

```bash
# All unit tests
cargo test --lib

# Specific test suite
cargo test v2::tests::integrated_api
cargo test v2::tests::legacy
cargo test v2::tests_right_panel

# Generate examples
cargo run --example test_multilang --release

# Performance benchmark
cargo run --example benchmark --release
```

### Test Output
- All test files write to: `tests/output/` (340 KB, 15 files)
- Benchmark uses temp directory: `/tmp/finanzbericht_benchmark/` (auto-cleanup)

---

## 🚀 Performance

### Benchmark Results (Release Build, 12-core CPU)

#### 100 Files
| Threads | Time | Files/sec | Speedup |
|---------|------|-----------|---------|
| 1 | 379 ms | 263.71 | 1.00x |
| 4 | 87 ms | 1149 | 4.36x |
| 8 | 68 ms | 1464 | 5.55x |
| 16 | 65 ms | 1535 | 5.82x |

#### 1000 Files
| Threads | Time | Files/sec | Speedup |
|---------|------|-----------|---------|
| 1 | 3.86s | 258.76 | 1.00x |
| 4 | 1.04s | 958 | 3.70x |
| 8 | 813 ms | 1231 | 4.76x |
| 16 | 714 ms | 1401 | 5.42x |

**Optimal Configuration:** 8 threads (good efficiency, minimal overhead)

---

## 📦 Dependencies

Key dependencies (see Cargo.toml):
- **rust_xlsxwriter** - Excel XLSX generation
- **serde** - Serialization
- **lazy_static** - Static initialization

---

## 🔍 Key Modules Explained

### 1. **api.rs** - Public API
- `ApiKey` enum - All configurable cell values
- `ReportValues` struct - Value container
- Methods: `set()`, `set_position_row()`, `set_header_input()`, etc.

### 2. **registry.rs** - Cell Registry
- `CellRegistry` - Central cell & formula management
- Topological sort for formula evaluation
- `computed` HashMap for caching results

### 3. **definitions.rs** - Formula Definitions
- `register_formula_cells()` - Register all static formulas
- Helper functions:
  - `register_text_lookup()` - VLOOKUP formulas
  - `register_iferror_division()` - Division with error handling
  - `register_sumproduct_round()` - SUMPRODUCT formulas
  - etc.

### 4. **writer.rs** - Main Write Function
- `write_report_v2()` - Main entry point
- Calls section writers
- Calls `write_cells_from_registry()` for formulas

### 5. **body/** - Dynamic Positions
- `BodyConfig` - How many positions per category
- `BodyLayout` - Where rows go, which rows to format
- `register_body_formulas()` - Register position-specific formulas
- `write_body_structure()` - Write position data

### 6. **sections/** - Static Sections
- `header.rs` - Project info, currency, dates
- `table.rs` - Income table
- `panel.rs` - K/R, O/V, L-U columns
- `prebody.rs` - Expenses summary

---

## 🛠️ Development Workflow

### Setup
```bash
# Clone or navigate to project
cd /home/ardit/repos/Arbeit/KMW-FB_Rust/src_rust

# Setup with devenv (if available)
devenv up

# Or use standard Rust
cargo build --release
```

### Development
```bash
# Watch for changes
cargo watch -x build -x test

# Run tests
cargo test --lib

# Run specific test
cargo test v2::tests::integrated_api::test_realistic_medium_project

# Run examples
cargo run --example test_multilang --release
```

### Benchmarking
```bash
cargo run --example benchmark --release
```

### Cleanup
```bash
# Clean build artifacts
cargo clean

# Remove test outputs
rm -rf tests/output/*.xlsx

# Regenerate cache if needed
rm -rf .devenv/
rm -f Cargo.lock devenv.lock
cargo build --release
```

---

## 📝 Recent Changes (2026-01-20)

### Cleanup Completed
- ✅ **Deleted V1 code** (88 KB, 6 Rust files, 2,215 lines)
  - Completely replaced by V2
  - No longer needed
  
- ✅ **Removed old output directories**
  - `/output/` - demo files (76 KB)
  - `src/v2/Excel-Template/output/` - extracted outputs (50 KB)
  - These were stale artifacts
  
- ✅ **Cleaned development cache**
  - `.devenv/` directory (1.3 MB)
  - Auto-regenerates on next `devenv` command

- ✅ **Consolidated test outputs**
  - Central location: `tests/output/`
  - All 15 test artifacts in one place

### Project Restructuring
- Created this `PROJECT_STRUCTURE.md` for clarity
- Organized modules logically
- Documented data flow
- Added quick reference guide

**Total Cleanup:** ~1.6 MB freed

---

## 🎯 Quick Reference

### Important Files
| File | Purpose |
|------|---------|
| `src/lib.rs` | Library entry point |
| `src/v2/report/api.rs` | Public API (ApiKey, ReportValues) |
| `src/v2/report/writer.rs` | Main write_report_v2() function |
| `src/v2/report/registry.rs` | Central formula management |
| `src/v2/report/definitions.rs` | All formula registrations |
| `src/v2/report/body/layout.rs` | Dynamic body layout calculation |
| `src/v2/report/formats.rs` | Formatting definitions |
| `examples/benchmark.rs` | Performance benchmarking |
| `examples/test_multilang.rs` | Multilingual report generation |

### Important Functions
| Function | Location | Purpose |
|----------|----------|---------|
| `write_report_v2()` | writer.rs | Main report generation |
| `build_registry()` | definitions.rs | Create registry with all formulas |
| `register_formula_cells()` | definitions.rs | Register static formulas |
| `register_body_formulas()` | body/formulas.rs | Register position formulas |
| `BodyLayout::compute()` | body/layout.rs | Calculate body row layout |
| `evaluate_all()` | registry.rs | Evaluate formulas with topological sort |

### Important Structures
| Struct | Location | Purpose |
|--------|----------|---------|
| `ReportValues` | values.rs | Value container (API values) |
| `BodyConfig` | body/config.rs | Positions per category config |
| `BodyLayout` | body/layout.rs | Row calculation results |
| `CellRegistry` | registry.rs | Central formula registry |
| `ReportStyles` | formats.rs | Format definitions |

---

## 📞 Support

### Running into issues?
1. Check `PROJECT_STRUCTURE.md` (this file) for module overview
2. Review test files for usage examples
3. Check `examples/` for working code
4. Run tests: `cargo test --lib`

### File locations
- Source code: `src/v2/`
- Tests: `src/v2/tests/` + `src/v2/tests_right_panel.rs`
- Examples: `examples/`
- Test output: `tests/output/`

---

**Last Updated:** 2026-01-20  
**Maintainer:** Development Team  
**Status:** ✅ Production Ready
