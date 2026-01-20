# Developer Guide - KMW-FB_Rust

Quick reference for developers working with the Financial Report Generator.

---

## 🚀 Getting Started

### First Time Setup

```bash
# Navigate to project
cd /home/ardit/repos/Arbeit/KMW-FB_Rust/src_rust

# Option 1: With DevEnv (NixOS)
devenv up
devenv shell

# Option 2: Standard Rust
rustup update
rustup default stable
```

### Verify Setup

```bash
cargo --version
cargo test --lib 2>&1 | grep "test result"
```

Expected output:
```
test result: ok. 92 passed; 0 failed; 0 ignored; 0 measured
```

---

## 📚 Key Concepts

### 1. ReportValues - API Container

**Where:** `src/v2/report/values.rs`

```rust
let mut values = ReportValues::new();

// Set static values
values.set(ApiKey::Language, "deutsch");
values.set(ApiKey::ProjectNumber, "KMW-2024-001");

// Set position data
values.set_position_row(
    category: 1,      // Budget category (1-5)
    position: 1,      // Position number within category
    description: "Personnel",
    approved: 50000.0,
    reported: 25000.0,
    approved_total: 25000.0,
    remark: ""
);

// Set header-input values (for categories 6-8)
values.set_header_input(
    category: 6,
    approved: 8000.0,
    reported: 4000.0,
    total: 4000.0,
    remark: ""
);
```

**Available ApiKeys:**
- `Language`, `Currency`, `ProjectNumber`, `ProjectTitle`
- `ProjectStart`, `ProjectEnd`, `ReportStart`, `ReportEnd`
- `ApprovedBudget(0-4)`, `IncomeReportPeriod(0-4)`, `IncomeTotal(0-4)`
- `LeftDate(0-17)`, `LeftAmountEuro(0-17)`, `LeftAmountLocal(0-17)`
- `RightDate(0-17)`, `RightAmountEuro(0-17)`, `RightAmountLocal(0-17)`
- Position methods: `set_position_row()`, `set_header_input()`
- Footer values: `set_footer_bank()`, `set_footer_kasse()`, `set_footer_sonstiges()`

### 2. BodyConfig - Position Configuration

**Where:** `src/v2/report/body/config.rs`

```rust
let body_config = BodyConfig::new()
    .with_positions(1, 10)  // Category 1: 10 positions
    .with_positions(2, 15)  // Category 2: 15 positions
    .with_positions(3, 8)   // Category 3: 8 positions
    .with_positions(4, 5)   // Category 4: 5 positions
    .with_positions(5, 3)   // Category 5: 3 positions
    .with_positions(6, 0)   // Category 6: header-input mode (0 = single row)
    .with_positions(7, 0)   // Category 7: header-input mode
    .with_positions(8, 0);  // Category 8: header-input mode
```

**Categories:**
- 1-5: Budget categories (can have multiple positions)
- 6-8: Special categories
  - `positions = 0` → Header-input mode (single data row)
  - `positions > 0` → Multiple positions (rarely used)

### 3. CellRegistry - Formula Management

**Where:** `src/v2/report/registry.rs`

The registry is the **single source of truth** for all formulas:

```rust
// In your code, formulas are automatically registered via:
let mut registry = build_registry(&values);  // from definitions.rs

// Registry operations
registry.is_formula(CellAddr::new(1, 1));      // Check if cell has formula
registry.register_formula(addr, formula);      // Register new formula
let computed = registry.evaluate_all(&values); // Evaluate all formulas
```

**Key Property:** Uses topological sorting to evaluate formulas in correct order.

### 4. BodyLayout - Position Layout Calculation

**Where:** `src/v2/report/body/layout.rs`

```rust
let layout = BodyLayout::compute(&body_config)?;

// Returns which rows to format and where positions go
// Internally calculates:
// - Header row position
// - Position data row positions
// - Sum rows
// - Footer rows
// - Total rows
```

---

## 🔄 Data Flow - Complete Example

### Step 1: Create Values & Config

```rust
use kmw_fb_rust::v2::report::{ApiKey, ReportValues, BodyConfig};

let mut values = ReportValues::new()
    .with_language("deutsch")
    .with_currency("EUR");

// Add position
values.set_position_row(1, 1, "Personnel", 10000.0, 5000.0, 5000.0, "");

let body_config = BodyConfig::new().with_positions(1, 1);
```

### Step 2: Create Worksheet & Styles

```rust
use rust_xlsxwriter::Workbook;
use kmw_fb_rust::v2::report::ReportStyles;

let mut workbook = Workbook::new();
let ws = workbook.add_worksheet();
let styles = ReportStyles::new();
```

### Step 3: Write Report

```rust
use kmw_fb_rust::v2::report::write_report_v2_with_body;

write_report_v2_with_body(ws, &styles, "_de", &values, &body_config)?;
ws.protect();
workbook.save("report.xlsx")?;
```

**What happens internally:**
1. `build_registry()` - Creates registry with all formulas
2. `write_header_section()` - Writes static header data
3. `write_table_section()` - Writes income table
4. `write_panel_section()` - Writes right panel
5. `write_prebody_section()` - Writes expense categories
6. `BodyLayout::compute()` - Calculates position row layout
7. `write_body_structure()` - Writes position header rows
8. `registry.evaluate_all()` - Evaluates all formulas with topological sort
9. `write_cells_from_registry()` - Writes formulas + formats

---

## 🧪 Testing Patterns

### Write a New Test

**Location:** `src/v2/tests/my_test.rs`

```rust
#[cfg(test)]
mod my_tests {
    use crate::v2::report::*;
    use rust_xlsxwriter::{Format, Workbook};
    
    #[test]
    fn test_my_feature() {
        // 1. Setup
        let mut workbook = Workbook::new();
        let ws = workbook.add_worksheet();
        
        // 2. Prepare data
        let mut values = ReportValues::new();
        values.set(ApiKey::Language, "deutsch");
        values.set_position_row(1, 1, "Test", 1000.0, 500.0, 500.0, "");
        
        let body_config = BodyConfig::new().with_positions(1, 1);
        
        // 3. Execute
        let styles = ReportStyles::new();
        write_report_v2_with_body(ws, &styles, "_de", &values, &body_config)
            .expect("Failed to write report");
        
        // 4. Assert (if needed)
        assert_eq!(values.get(ApiKey::Language), Some("deutsch"));
        
        // 5. Save for inspection
        ws.protect();
        workbook.save("tests/output/test_my_feature.xlsx")
            .expect("Failed to save");
    }
}
```

### Register in test module

**File:** `src/v2/tests/mod.rs`

```rust
mod body_config;
mod body_layout;
mod integrated_api;
mod legacy;
mod my_test;        // Add this
mod position_api;
```

### Run your test

```bash
cargo test my_tests::test_my_feature
```

---

## 🔍 Code Organization Decisions

### Why Registry?
- **Single Source of Truth** for all formulas
- **Automatic Topological Sorting** for correct evaluation order
- **Easy to Add Formulas** - just call `register_formula()`
- **Separation of Concerns** - formulas separate from layout

### Why Sections?
- **Modular Design** - each section is independent
- **Easy to Extend** - add new section without touching others
- **Clear Responsibility** - header knows header, body knows body
- **Testable** - can test each section individually

### Why BodyLayout?
- **Flexible Positions** - any number of positions per category
- **Complex Row Calculations** - consolidates complex logic
- **Reusable** - used for both reading and writing
- **Tested** - has 9 unit tests

---

## 💡 Common Tasks

### Add a New API Value

**1. Add to ApiKey enum** → `src/v2/report/api.rs`

```rust
pub enum ApiKey {
    // ... existing ...
    MyNewValue,  // Add this
}
```

**2. Handle in ReportValues** → `src/v2/report/values.rs`

```rust
impl ReportValues {
    pub fn set_my_value(&mut self, value: impl Into<String>) -> &mut Self {
        self.values.insert(ApiKey::MyNewValue, value.into());
        self
    }
    
    pub fn get_my_value(&self) -> Option<&str> {
        self.values.get(&ApiKey::MyNewValue).map(|v| v.as_str())
    }
}
```

**3. Register formula** → `src/v2/report/definitions.rs`

```rust
pub fn register_formula_cells(registry: &mut CellRegistry<...>, ...) {
    // If MyNewValue should have a formula:
    registry.register_formula(
        CellAddr::new(row, col),
        Formula::Value(lookup_text(language, index))
    );
}
```

### Add a New Formula

**1. Define formula** → `src/v2/report/formulas.rs`

```rust
pub fn my_formula(param: f64) -> Formula {
    Formula::String(format!("=SUM(A1:A5)*{}", param))
}
```

**2. Register it** → `src/v2/report/definitions.rs`

```rust
registry.register_formula(
    CellAddr::new(10, 5),
    my_formula(parameter)
);
```

**3. Add test** → `src/v2/tests/legacy.rs`

```rust
#[test]
fn test_my_formula() {
    let registry = build_registry(&values);
    assert!(registry.is_formula(CellAddr::new(10, 5)));
}
```

### Add Multi-Position Support

**1. Update BodyConfig** → `src/v2/report/body/config.rs`

```rust
let config = BodyConfig::new()
    .with_positions(1, 50)  // Support up to 50 positions
    .with_positions(2, 30);
```

**2. Formulas adapt automatically** via `body/formulas.rs`

The system automatically generates formulas for all positions.

---

## 🐛 Debugging Tips

### Enable Test Output

```bash
cargo test my_test -- --nocapture
```

### Inspect Generated Excel

All tests write to `tests/output/`. Open the .xlsx file in Excel:
```bash
open tests/output/test_realistic_medium_project.xlsx
```

### Check Formula in Cell

In Excel: Click cell → View formula bar → See formula

### Trace Evaluation Order

Add debug output in `registry.rs`:
```rust
println!("Evaluating: {:?}", addr);
```

Then run:
```bash
cargo test my_test -- --nocapture 2>&1 | grep "Evaluating"
```

### Check Value Cache

In `registry.rs`, inspect `self.computed`:
```rust
println!("Cached values: {:?}", self.computed);
```

---

## 📊 Performance Profiling

### Run Benchmark

```bash
cargo run --example benchmark --release
```

Output shows:
- Files per second
- Speedup with threads
- Optimal thread count

### Profile Memory

```bash
# Using cargo-flamegraph (if installed)
cargo install flamegraph
cargo flamegraph --example benchmark --release
```

### Check Build Time

```bash
cargo build --release -v 2>&1 | grep "Compiling"
```

---

## 🔧 Useful Commands

```bash
# Build & test
cargo build --release
cargo test --lib
cargo test --lib v2::tests::integrated_api

# Run examples
cargo run --example test_multilang --release
cargo run --example benchmark --release

# Development workflow
cargo watch -x build -x test
cargo clippy
cargo fmt

# Clean & rebuild
cargo clean
cargo build --release

# Documentation
cargo doc --open
```

---

## 📁 File Navigation Quick Guide

**Want to...**

| Task | File(s) |
|------|---------|
| Add API value | `values.rs`, `api.rs` |
| Add formula | `formulas.rs`, `definitions.rs` |
| Modify header | `sections/header.rs` |
| Modify body | `body/layout.rs`, `body/writer.rs` |
| Write test | `tests/*.rs`, `tests_right_panel.rs` |
| Fix styling | `formats.rs` |
| Debug formula | `registry.rs`, `dynamic.rs` |
| Add language | `lang/data.rs` |

---

## ✅ Code Style

### Naming Conventions
- Functions: `snake_case`
- Types/Structs: `PascalCase`
- Constants: `SCREAMING_SNAKE_CASE`
- Variables: `snake_case`

### Organization
- One concept per file
- Max 500 lines per file (usually)
- Public API at top of file
- Tests at bottom or separate file

### Comments
- Explain **why**, not **what**
- Doc comments for public APIs
- Avoid obvious comments

---

## 🚀 Deployment Checklist

Before shipping:

- [ ] `cargo test --lib` passes (92/92 tests)
- [ ] `cargo build --release` succeeds
- [ ] `cargo clippy` has no warnings
- [ ] `cargo fmt` is applied
- [ ] Benchmark runs: `cargo run --example benchmark --release`
- [ ] Examples generate correctly: `cargo run --example test_multilang --release`
- [ ] Documentation is up to date

---

## 📞 Getting Help

1. **Architecture Questions** → Read `PROJECT_STRUCTURE.md`
2. **API Usage** → Check `examples/test_multilang.rs`
3. **Test Examples** → Look at `src/v2/tests/`
4. **Performance** → Run `cargo run --example benchmark --release`

---

**Last Updated:** 2026-01-20  
**For Issues:** Check existing code, tests, and examples first
