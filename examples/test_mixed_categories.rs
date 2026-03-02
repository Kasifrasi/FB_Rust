//! Test: Mixed category modes for categories 6–8
//!
//! Demonstriert drei verschiedene Kombinationen für die Kostenkategorien 6, 7 und 8:
//!
//! 1. **Standard** — Kategorien 1–5 als Multi-Row, Kategorien 6–8 als Header-Input
//!    (klassische Verwendung: aggregierte Einzelwerte ohne Positions-Zeilen)
//!
//! 2. **Multi-Row** — Kategorien 6–8 *ebenfalls* mit mehreren Zeilen via `add_position()`
//!    (z.B. wenn Kategorie 6 mehrere Einzel-Buchungen ausweisen soll)
//!
//! 3. **Gemischt** — Kategorien 1–4 Multi-Row, Kategorie 5 Header-Input,
//!    Kategorien 6–8 Multi-Row mit Skip-Zeilen
//!
//! **Builder-Pattern** — kein `Some()`, kein `.to_string()`, kein `?` beim Bauen.

use fb_rust::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    std::fs::create_dir_all("examples/output")?;

    println!("Testing mixed category modes for categories 6-8...\n");

    // -------------------------------------------------------------------------
    // Test 1: Kategorien 6-8 als klassische Header-Inputs (aggregierte Werte)
    // -------------------------------------------------------------------------
    {
        println!("1. Categories 6-8 as header-inputs (aggregated single values)...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("MIX-001")
                    .project_title("Gemischte Kategorien Test")
                    .project_start("01.01.2025")
                    .project_end("31.12.2027")
                    .report_start("01.01.2025")
                    .report_end("30.06.2025")
                    .build(),
            )
            .body(
                ReportBody::builder()
                    .kmw_mittel(
                        TableEntry::builder()
                            .approved_budget(80_000.0)
                            .income_report(50_000.0)
                            .income_total(50_000.0)
                            .reason("1. Rate")
                            .build(),
                    )
                    // Kategorien 1-5: Multi-Row Positionen
                    .add_position(1, PositionEntry::builder().description("Projektleitung").approved(18_000.0).income_report(9_000.0).income_total(9_000.0).build())
                    .add_position(1, PositionEntry::builder().description("Buchhaltung").approved(5_000.0).build())
                    .add_position(2, PositionEntry::builder().description("Reisekosten").approved(3_000.0).income_report(1_500.0).build())
                    .add_position(3, PositionEntry::builder().description("Sachkosten").approved(2_000.0).build())
                    // Kategorien 6-8: Header-Inputs (klassisch — aggregierter Einzelwert)
                    .set_header_input(6, PositionEntry::builder().approved(3_000.0).remark("Verwaltung pauschal").build())
                    .set_header_input(7, PositionEntry::builder().approved(1_500.0).remark("Reservefonds").build())
                    .set_header_input(8, PositionEntry::builder().approved(800.0).remark("Sonstiges").build())
                    .build(),
            )
            .footer(
                ReportFooter::builder()
                    .bank(8_500.0)
                    .kasse(1_200.0)
                    .sonstiges(300.0)
                    .build(),
            )
            .build();
        config.write_to("examples/output/test_mixed_cat_header_inputs.xlsx")?;
        println!("   OK: examples/output/test_mixed_cat_header_inputs.xlsx");
        println!("   Categories 6/7/8 → header-input (no individual position rows)");
    }

    // -------------------------------------------------------------------------
    // Test 2: Kategorien 6-8 als Multi-Row (mehrere Einzel-Positionen)
    // -------------------------------------------------------------------------
    {
        println!("\n2. Categories 6-8 as multi-row positions (multiple entries each)...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::English)
                    .currency(Currency::EUR)
                    .project_number("MIX-002")
                    .project_title("Multi-Row Categories 6-8")
                    .project_start("01.01.2025")
                    .project_end("31.12.2027")
                    .report_start("01.01.2025")
                    .report_end("30.06.2025")
                    .build(),
            )
            .body(
                ReportBody::builder()
                    .kmw_mittel(TableEntry::builder().approved_budget(60_000.0).income_report(30_000.0).income_total(30_000.0).build())
                    .eigenmittel(TableEntry::builder().approved_budget(10_000.0).income_report(5_000.0).income_total(5_000.0).build())
                    // Kategorien 1-5: Multi-Row
                    .add_position(1, PositionEntry::builder().description("Staff costs").approved(20_000.0).income_report(10_000.0).income_total(10_000.0).build())
                    .add_position(2, PositionEntry::builder().description("Travel").approved(4_000.0).income_report(2_000.0).build())
                    .add_position(5, PositionEntry::builder().description("Equipment").approved(3_000.0).build())
                    // Kategorie 6: Multi-Row (z.B. mehrere Verwaltungsposten)
                    .add_position(6, PositionEntry::builder().description("Office rent").approved(1_200.0).income_report(600.0).income_total(600.0).build())
                    .add_position(6, PositionEntry::builder().description("Phone & internet").approved(360.0).income_report(180.0).income_total(180.0).build())
                    .add_position(6, PositionEntry::builder().description("Accounting software").approved(240.0).build())
                    // Kategorie 7: Multi-Row mit Skip-Zeile
                    .add_position(7, PositionEntry::builder().description("Contingency fund").approved(2_000.0).build())
                    .skip_position_row(7)
                    .add_position(7, PositionEntry::builder().description("Emergency reserve").approved(1_000.0).build())
                    // Kategorie 8: Einzelne Position
                    .add_position(8, PositionEntry::builder().description("Evaluation costs").approved(500.0).income_report(250.0).build())
                    .build(),
            )
            .footer(
                ReportFooter::builder()
                    .bank(12_000.0)
                    .kasse(800.0)
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .hide_columns_qv(true)
                    .build(),
            )
            .build();
        config.write_to("examples/output/test_mixed_cat_multi_row.xlsx")?;
        println!("   OK: examples/output/test_mixed_cat_multi_row.xlsx");
        println!("   Category 6 → 3 rows | Category 7 → 2 rows + 1 skip | Category 8 → 1 row");
    }

    // -------------------------------------------------------------------------
    // Test 3: Gemischt — unterschiedliche Modi pro Kategorie + version-Feld
    // -------------------------------------------------------------------------
    {
        println!("\n3. Mixed modes: some categories as multi-row, others as header-input...");
        let config = ReportConfig::builder()
            .header(
                ReportHeader::builder()
                    .language(Language::Deutsch)
                    .currency(Currency::EUR)
                    .project_number("MIX-003")
                    .project_title("Gemischte Modi")
                    .project_start("01.01.2025")
                    .project_end("31.12.2027")
                    .report_start("01.01.2025")
                    .report_end("30.06.2025")
                    .version("v2025-MIX")  // wird in Zelle B2 geschrieben
                    .build(),
            )
            .body(
                ReportBody::builder()
                    .kmw_mittel(TableEntry::builder().approved_budget(100_000.0).income_report(40_000.0).income_total(40_000.0).reason("1. Tranche").build())
                    .saldovortrag(TableEntry::builder().approved_budget(5_000.0).build())
                    // Kategorien 1-4: Multi-Row Positionen
                    .add_position(1, PositionEntry::builder().description("Koordinator/in").approved(24_000.0).income_report(12_000.0).income_total(12_000.0).remark("50%").build())
                    .add_position(1, PositionEntry::builder().description("Buchhalter/in").approved(6_000.0).income_report(3_000.0).income_total(3_000.0).build())
                    .skip_position_row(1)
                    .add_position(1, PositionEntry::builder().description("Freelancer").approved(4_000.0).build())
                    .add_position(2, PositionEntry::builder().description("Inlandsreisen").approved(2_000.0).income_report(800.0).build())
                    .add_position(2, PositionEntry::builder().description("Auslandsreisen").approved(5_000.0).income_report(2_500.0).income_total(2_500.0).build())
                    .add_position(3, PositionEntry::builder().description("Büromaterial").approved(500.0).build())
                    .add_position(4, PositionEntry::builder().description("IT-Ausstattung").approved(3_000.0).income_report(3_000.0).income_total(3_000.0).build())
                    // Kategorie 5: Header-Input (Gesamtbetrag, keine Einzel-Zeilen)
                    .set_header_input(5, PositionEntry::builder().approved(8_000.0).remark("Direktkosten gesamt").build())
                    // Kategorien 6-8: Multi-Row (detaillierte Indirektkosten)
                    .add_position(6, PositionEntry::builder().description("Büromiete").approved(3_600.0).income_report(1_800.0).income_total(1_800.0).build())
                    .add_position(6, PositionEntry::builder().description("Telekommunikation").approved(480.0).income_report(240.0).income_total(240.0).build())
                    .add_position(7, PositionEntry::builder().description("Rücklagen").approved(3_000.0).build())
                    .add_position(8, PositionEntry::builder().description("Evaluation").approved(2_000.0).income_report(1_000.0).build())
                    .add_position(8, PositionEntry::builder().description("Audit").approved(1_500.0).build())
                    .build(),
            )
            .footer(
                ReportFooter::builder()
                    .bank(15_000.0)
                    .kasse(2_500.0)
                    .sonstiges(500.0)
                    .build(),
            )
            .options(
                ReportOptions::builder()
                    .sheet_password("geheim")
                    .hide_columns_qv(true)
                    .hide_language_sheet(true)
                    .build(),
            )
            .build();
        config.write_to("examples/output/test_mixed_cat_gemischt.xlsx")?;
        println!("   OK: examples/output/test_mixed_cat_gemischt.xlsx");
        println!("   Category 1 → 3 rows + 1 skip | Category 5 → header-input");
        println!("   Category 6 → 2 rows | Category 7 → 1 row | Category 8 → 2 rows");
        println!("   B2 cell → \"v2025-MIX\"");
    }

    println!("\nAll tests completed successfully!");
    Ok(())
}
