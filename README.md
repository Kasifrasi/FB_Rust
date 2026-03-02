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

`ReportConfig` is the main entry point — composed of four domain sub-structs:
`ReportHeader`, `ReportBody`, `ReportFooter`, and `ReportOptions`.

All structs use a hand-written fluent builder with infallible `.build()` — no `?` required:

```rust
//! Test: Finanzbericht mit ALLEN Feldern gefüllt (E = F)
//!
//! Generiert einen vollständigen Finanzbericht, bei dem:
//! - Alle Header-Felder gesetzt sind
//! - Alle 5 Einnahmen-Zeilen gefüllt sind
//! - Alle 8 Kostenkategorien (5 mit Positionen, 3 Header-Eingabe) gefüllt sind
//! - Alle 36 Belege im Panel gefüllt sind
//! - Alle Footer-Felder (Bank, Kasse, Sonstiges) gesetzt sind
//! - Spalte E (Berichtsperiode) = Spalte F (Kumulativ) — wie im 1. Berichtszeitraum
//!
//! **Builder-Pattern Showcase** — kein `Some()`, kein `.to_string()`, kein `?` beim Bauen.

use fb_rust::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle vollständigen Finanzbericht (E = F)...\n");

    let config = ReportConfig::builder()
        .header(
            ReportHeader::builder()
                .language(Language::Deutsch)
                .currency(Currency::EUR)
                .project_number("PROJ-2024-001")
                .project_title("Bildungsprojekt Ostafrika")
                .project_start("01.01.2024")
                .project_end("31.12.2026")
                .report_start("01.01.2024")
                .report_end("30.06.2024")
                .version("v2025-1")
                .build(),
        )
        .body(
            ReportBody::builder()
                // ============================================================
                // Einnahmen-Tabelle (5 benannte Zeilen)
                // ============================================================
                .kmw_mittel(   TableEntry::builder().approved_budget(80_000.0).income_report(50_000.0).income_total(50_000.0).reason("1. Rate erhalten").build())
                .eigenmittel(  TableEntry::builder().approved_budget(15_000.0).income_report( 7_500.0).income_total( 7_500.0).reason("Eigenanteil eingebracht").build())
                .drittmittel(  TableEntry::builder().approved_budget(10_000.0).income_report( 5_000.0).income_total( 5_000.0).reason("Stiftung XY").build())
                .saldovortrag( TableEntry::builder().approved_budget( 2_500.0).income_report( 1_500.0).income_total( 1_500.0).reason("Spenden").build())
                .zinsertraege( TableEntry::builder().approved_budget(   500.0).income_report(   400.0).income_total(   400.0).reason("Tagesgeldkonto").build())
                // ============================================================
                // Linkes Panel (Belege 1–18)
                // ============================================================
                .add_left_panel_entry(PanelEntry::builder().date("15.01.2024").amount_euro( 9_000.0).amount_local( 9_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("15.01.2024").amount_euro( 6_000.0).amount_local( 6_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("31.01.2024").amount_euro( 4_000.0).amount_local( 4_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("31.01.2024").amount_euro( 3_000.0).amount_local( 3_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("15.02.2024").amount_euro( 2_000.0).amount_local( 2_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("28.02.2024").amount_euro( 2_500.0).amount_local( 2_500.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("01.03.2024").amount_euro( 1_500.0).amount_local( 1_500.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("15.03.2024").amount_euro(   750.0).amount_local(   750.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("31.03.2024").amount_euro( 1_000.0).amount_local( 1_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("01.04.2024").amount_euro( 2_500.0).amount_local( 2_500.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("15.04.2024").amount_euro( 1_800.0).amount_local( 1_800.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("30.04.2024").amount_euro(   600.0).amount_local(   600.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("01.05.2024").amount_euro( 6_000.0).amount_local( 6_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("15.05.2024").amount_euro( 1_500.0).amount_local( 1_500.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("31.05.2024").amount_euro( 3_000.0).amount_local( 3_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("01.06.2024").amount_euro( 2_000.0).amount_local( 2_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("15.06.2024").amount_euro( 1_000.0).amount_local( 1_000.0).build())
                .add_left_panel_entry(PanelEntry::builder().date("30.06.2024").amount_euro( 4_000.0).amount_local( 4_000.0).build())
                // ============================================================
                // Rechtes Panel (Belege 19–36)
                // ============================================================
                .add_right_panel_entry(PanelEntry::builder().date("05.01.2024").amount_euro( 1_500.0).amount_local( 1_500.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("10.01.2024").amount_euro(   750.0).amount_local(   750.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("20.01.2024").amount_euro(   500.0).amount_local(   500.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("05.02.2024").amount_euro( 2_000.0).amount_local( 2_000.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("15.02.2024").amount_euro( 1_200.0).amount_local( 1_200.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("28.02.2024").amount_euro(   800.0).amount_local(   800.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("10.03.2024").amount_euro(   950.0).amount_local(   950.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("20.03.2024").amount_euro( 1_100.0).amount_local( 1_100.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("01.04.2024").amount_euro( 1_800.0).amount_local( 1_800.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("15.04.2024").amount_euro(   650.0).amount_local(   650.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("30.04.2024").amount_euro( 2_200.0).amount_local( 2_200.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("10.05.2024").amount_euro( 1_350.0).amount_local( 1_350.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("20.05.2024").amount_euro(   900.0).amount_local(   900.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("01.06.2024").amount_euro( 1_600.0).amount_local( 1_600.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("15.06.2024").amount_euro( 2_100.0).amount_local( 2_100.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("25.06.2024").amount_euro(   750.0).amount_local(   750.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("28.06.2024").amount_euro( 1_400.0).amount_local( 1_400.0).build())
                .add_right_panel_entry(PanelEntry::builder().date("30.06.2024").amount_euro(   850.0).amount_local(   850.0).build())
                // ============================================================
                // Kostenpositionen (Kategorien 1–5): add_position(kategorie, eintrag)
                // Keine category/position Felder nötig — Reihenfolge = Position
                // ============================================================
                // Kategorie 1: Bauausgaben — 5 Positionen
                .add_position(1, PositionEntry::builder().description("Projektleitung").approved(18_000.0).income_report( 9_000.0).income_total( 9_000.0).remark("6 Monate").build())
                .add_position(1, PositionEntry::builder().description("Buchhaltung").approved(12_000.0).income_report( 6_000.0).income_total( 6_000.0).remark("Teilzeit").build())
                .add_position(1, PositionEntry::builder().description("Lokale Fachkräfte").approved( 8_000.0).income_report( 4_000.0).income_total( 4_000.0).remark("2 Mitarbeiter").build())
                .add_position(1, PositionEntry::builder().description("Sozialabgaben").approved( 6_000.0).income_report( 3_000.0).income_total( 3_000.0).remark("AG-Anteile").build())
                .add_position(1, PositionEntry::builder().description("Honorare Trainer").approved( 4_000.0).income_report( 2_000.0).income_total( 2_000.0).remark("3 Workshops").build())
                // Kategorie 2: Investitionen — 3 Positionen
                .add_position(2, PositionEntry::builder().description("Flüge International").approved( 5_000.0).income_report( 2_500.0).income_total( 2_500.0).remark("2 Dienstreisen").build())
                .add_position(2, PositionEntry::builder().description("Unterkunft/Verpflegung").approved( 3_000.0).income_report( 1_500.0).income_total( 1_500.0).remark("Tagessätze").build())
                .add_position(2, PositionEntry::builder().description("Lokale Transporte").approved( 1_500.0).income_report(   750.0).income_total(   750.0).remark("Projektfahrten").build())
                // Kategorie 3: Personalausgaben — 4 Positionen
                .add_position(3, PositionEntry::builder().description("Büromaterial").approved( 2_000.0).income_report( 1_000.0).income_total( 1_000.0).remark("Verbrauchsmaterial").build())
                .add_position(3, PositionEntry::builder().description("Lehrmaterialien").approved( 5_000.0).income_report( 2_500.0).income_total( 2_500.0).remark("Schulbücher + Hefte").build())
                .add_position(3, PositionEntry::builder().description("Miete Schulungsraum").approved( 3_600.0).income_report( 1_800.0).income_total( 1_800.0).remark("6 Monate").build())
                .add_position(3, PositionEntry::builder().description("Kommunikation").approved( 1_200.0).income_report(   600.0).income_total(   600.0).remark("Internet + Telefon").build())
                // Kategorie 4: Projektaktivitäten — 3 Positionen
                .add_position(4, PositionEntry::builder().description("Laptops").approved( 6_000.0).income_report( 6_000.0).income_total( 6_000.0).remark("4 Stück angeschafft").build())
                .add_position(4, PositionEntry::builder().description("Beamer").approved( 1_500.0).income_report( 1_500.0).income_total( 1_500.0).remark("1 Stück").build())
                .add_position(4, PositionEntry::builder().description("Mobiliar").approved( 3_000.0).income_report( 3_000.0).income_total( 3_000.0).remark("Tische + Stühle").build())
                // Kategorie 5: Projektverwaltung — 2 Positionen
                .add_position(5, PositionEntry::builder().description("Lehrerfortbildung").approved( 4_000.0).income_report( 2_000.0).income_total( 2_000.0).remark("2 Seminare").build())
                .add_position(5, PositionEntry::builder().description("Supervision").approved( 2_000.0).income_report( 1_000.0).income_total( 1_000.0).remark("Quartalstreffen").build())
                // ============================================================
                // Header-Input (Kategorien 6–8): set_header_input(kategorie, eintrag)
                // Einzelner Aggregatwert pro Kategorie — keine Einzelpositionen
                // ============================================================
                .set_header_input(6, PositionEntry::builder().approved( 8_000.0).income_report( 4_000.0).income_total( 4_000.0).remark("Verwaltungspauschale").build())
                .set_header_input(7, PositionEntry::builder().approved( 3_000.0).income_report( 1_500.0).income_total( 1_500.0).remark("Broschüren + Website").build())
                .set_header_input(8, PositionEntry::builder().approved( 1_200.0).income_report(   750.0).income_total(   750.0).remark("Bankgebühren").build())
                .build(),
        )
        .footer(
            ReportFooter::builder()
                .bank(8_500.0)
                .kasse(1_200.0)
                .sonstiges(300.0)
                .build(),
        )
        .options(
            ReportOptions::builder()
                .sheet_password("blatt_geheim")
                .workbook_password("wb_geheim")
                .hide_columns_qv(true)
                .hide_language_sheet(true)
                .build(),
        )
        .build();

    let output_path = "examples/output/test_all_fields.xlsx";
    config.write_to(output_path)?;

    println!("Finanzbericht erstellt: {}", output_path);
    println!();
    println!("Zusammenfassung:");
    println!("  Header:           8 Felder gesetzt");
    println!("  Einnahmen:        5 Zeilen x 4 Spalten (D, E, F, H)");
    println!("  Kostenpositionen: 17 Positionen + 3 Header-Eingaben");
    println!("  Belege:           36 Einträge (18 links + 18 rechts)");
    println!("  Footer:           Bank={}, Kasse={}, Sonstiges={}", 8500.0, 1200.0, 300.0);
    println!();
    println!("  Spalte E = Spalte F (1. Berichtszeitraum = Kumulativ)");

    Ok(())
}
```

> See [`examples/test_all_fields.rs`](examples/test_all_fields.rs) for a complete reference with all fields populated.
> See [`examples/test_mixed_categories.rs`](examples/test_mixed_categories.rs) for mixed category modes (header-input vs. multi-row for categories 6–8).

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
└── serde_integration.rs    Serde JSON roundtrip tests (28 tests, feature-gated)

examples/
├── test_all_fields.rs         All fields populated (complete reference example)
├── test_multilang.rs          5 languages
├── test_multilang_grouped.rs  Row grouping / collapsed sections
├── test_workbook_protection.rs  3 protection levels
├── test_parallel_protection.rs  Parallel protection benchmark (Rayon)
├── test_mixed_categories.rs    Mixed category modes (header-input vs. multi-row for 6–8)
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
cargo test                                             # Unit tests (94 tests)
cargo test --features serde --test serde_integration   # Serde JSON tests (28 tests)
cargo test --features serde                            # All tests (122 tests)
cargo deny check                                       # License and security audit
```

## License

Licensed under either of

- [MIT License](LICENSE-MIT)
- [Apache License, Version 2.0](LICENSE-2.0)

at your option.
