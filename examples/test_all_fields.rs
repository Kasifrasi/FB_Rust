//! Test: Finanzbericht mit ALLEN Feldern gefüllt (E = F)
//!
//! Generiert einen vollständigen Finanzbericht, bei dem:
//! - Alle Header-Felder gesetzt sind
//! - Alle 5 Einnahmen-Zeilen gefüllt sind
//! - Alle 8 Kostenkategorien (5 mit Positionen, 3 Header-Eingabe) gefüllt sind
//! - Alle 36 Belege im Right Panel gefüllt sind
//! - Alle Footer-Felder (Bank, Kasse, Sonstiges) gesetzt sind
//! - Spalte E (Berichtsperiode) = Spalte F (Kumulativ) — wie im 1. Berichtszeitraum
//!
//! **Builder-Pattern Showcase** — kein `Some()`, kein `.to_string()`.

use fb_rust::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle vollständigen Finanzbericht (E = F)...\n");

    let config = ReportConfigBuilder::default()
        .header(
            ReportHeaderBuilder::default()
                .language("deutsch")
                .currency("EUR")
                .project_number("PROJ-2024-001")
                .project_title("Bildungsprojekt Ostafrika")
                .project_start("01.01.2024")
                .project_end("31.12.2026")
                .report_start("01.01.2024")
                .report_end("30.06.2024")
                .build()?
        )
        .body(
            ReportBodyBuilder::default()
                // =============================================================
                // Einnahmen-Tabelle (Rows 15-19) — E = F
                // =============================================================
                .add_table_entry(TableEntryBuilder::default().index(0u8).approved_budget(80000.0).income_report(50000.0).income_total(50000.0).reason("1. Rate erhalten").build()?)
                .add_table_entry(TableEntryBuilder::default().index(1u8).approved_budget(15000.0).income_report(7500.0).income_total(7500.0).reason("Eigenanteil eingebracht").build()?)
                .add_table_entry(TableEntryBuilder::default().index(2u8).approved_budget(10000.0).income_report(5000.0).income_total(5000.0).reason("Stiftung XY").build()?)
                .add_table_entry(TableEntryBuilder::default().index(3u8).approved_budget(2500.0).income_report(1500.0).income_total(1500.0).reason("Spenden").build()?)
                .add_table_entry(TableEntryBuilder::default().index(4u8).approved_budget(500.0).income_report(400.0).income_total(400.0).reason("Tagesgeldkonto").build()?)
                // =============================================================
                // Linkes Panel (Belege 1-18)
                // =============================================================
                .add_left_panel_entry(PanelEntryBuilder::default().index(0u8).date("15.01.2024").amount_euro(9000.0).amount_local(9000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(1u8).date("15.01.2024").amount_euro(6000.0).amount_local(6000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(2u8).date("31.01.2024").amount_euro(4000.0).amount_local(4000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(3u8).date("31.01.2024").amount_euro(3000.0).amount_local(3000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(4u8).date("15.02.2024").amount_euro(2000.0).amount_local(2000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(5u8).date("28.02.2024").amount_euro(2500.0).amount_local(2500.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(6u8).date("01.03.2024").amount_euro(1500.0).amount_local(1500.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(7u8).date("15.03.2024").amount_euro(750.0).amount_local(750.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(8u8).date("31.03.2024").amount_euro(1000.0).amount_local(1000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(9u8).date("01.04.2024").amount_euro(2500.0).amount_local(2500.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(10u8).date("15.04.2024").amount_euro(1800.0).amount_local(1800.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(11u8).date("30.04.2024").amount_euro(600.0).amount_local(600.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(12u8).date("01.05.2024").amount_euro(6000.0).amount_local(6000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(13u8).date("15.05.2024").amount_euro(1500.0).amount_local(1500.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(14u8).date("31.05.2024").amount_euro(3000.0).amount_local(3000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(15u8).date("01.06.2024").amount_euro(2000.0).amount_local(2000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(16u8).date("15.06.2024").amount_euro(1000.0).amount_local(1000.0).build()?)
                .add_left_panel_entry(PanelEntryBuilder::default().index(17u8).date("30.06.2024").amount_euro(4000.0).amount_local(4000.0).build()?)
                // =============================================================
                // Rechtes Panel (Belege 19-36)
                // =============================================================
                .add_right_panel_entry(PanelEntryBuilder::default().index(0u8).date("05.01.2024").amount_euro(1500.0).amount_local(1500.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(1u8).date("10.01.2024").amount_euro(750.0).amount_local(750.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(2u8).date("20.01.2024").amount_euro(500.0).amount_local(500.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(3u8).date("05.02.2024").amount_euro(2000.0).amount_local(2000.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(4u8).date("15.02.2024").amount_euro(1200.0).amount_local(1200.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(5u8).date("28.02.2024").amount_euro(800.0).amount_local(800.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(6u8).date("10.03.2024").amount_euro(950.0).amount_local(950.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(7u8).date("20.03.2024").amount_euro(1100.0).amount_local(1100.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(8u8).date("01.04.2024").amount_euro(1800.0).amount_local(1800.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(9u8).date("15.04.2024").amount_euro(650.0).amount_local(650.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(10u8).date("30.04.2024").amount_euro(2200.0).amount_local(2200.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(11u8).date("10.05.2024").amount_euro(1350.0).amount_local(1350.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(12u8).date("20.05.2024").amount_euro(900.0).amount_local(900.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(13u8).date("01.06.2024").amount_euro(1600.0).amount_local(1600.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(14u8).date("15.06.2024").amount_euro(2100.0).amount_local(2100.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(15u8).date("25.06.2024").amount_euro(750.0).amount_local(750.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(16u8).date("28.06.2024").amount_euro(1400.0).amount_local(1400.0).build()?)
                .add_right_panel_entry(PanelEntryBuilder::default().index(17u8).date("30.06.2024").amount_euro(850.0).amount_local(850.0).build()?)
                // =============================================================
                // Kostenpositionen — E = F überall
                // =============================================================
                // Kategorie 1: Bauausgaben (5 Positionen) — Summe D=48000, E=F=24000
                .add_position(PositionEntryBuilder::default().category(1u8).position(1u16).description("Projektleitung").approved(18000.0).income_report(9000.0).income_total(9000.0).remark("6 Monate").build()?)
                .add_position(PositionEntryBuilder::default().category(1u8).position(2u16).description("Buchhaltung").approved(12000.0).income_report(6000.0).income_total(6000.0).remark("Teilzeit").build()?)
                .add_position(PositionEntryBuilder::default().category(1u8).position(3u16).description("Lokale Fachkräfte").approved(8000.0).income_report(4000.0).income_total(4000.0).remark("2 Mitarbeiter").build()?)
                .add_position(PositionEntryBuilder::default().category(1u8).position(4u16).description("Sozialabgaben").approved(6000.0).income_report(3000.0).income_total(3000.0).remark("AG-Anteile").build()?)
                .add_position(PositionEntryBuilder::default().category(1u8).position(5u16).description("Honorare Trainer").approved(4000.0).income_report(2000.0).income_total(2000.0).remark("3 Workshops").build()?)
                // Kategorie 2: Investitionen (3 Positionen) — Summe D=9500, E=F=4750
                .add_position(PositionEntryBuilder::default().category(2u8).position(1u16).description("Flüge International").approved(5000.0).income_report(2500.0).income_total(2500.0).remark("2 Dienstreisen").build()?)
                .add_position(PositionEntryBuilder::default().category(2u8).position(2u16).description("Unterkunft/Verpflegung").approved(3000.0).income_report(1500.0).income_total(1500.0).remark("Tagessätze").build()?)
                .add_position(PositionEntryBuilder::default().category(2u8).position(3u16).description("Lokale Transporte").approved(1500.0).income_report(750.0).income_total(750.0).remark("Projektfahrten").build()?)
                // Kategorie 3: Personalausgaben (4 Positionen) — Summe D=11800, E=F=5900
                .add_position(PositionEntryBuilder::default().category(3u8).position(1u16).description("Büromaterial").approved(2000.0).income_report(1000.0).income_total(1000.0).remark("Verbrauchsmaterial").build()?)
                .add_position(PositionEntryBuilder::default().category(3u8).position(2u16).description("Lehrmaterialien").approved(5000.0).income_report(2500.0).income_total(2500.0).remark("Schulbücher + Hefte").build()?)
                .add_position(PositionEntryBuilder::default().category(3u8).position(3u16).description("Miete Schulungsraum").approved(3600.0).income_report(1800.0).income_total(1800.0).remark("6 Monate").build()?)
                .add_position(PositionEntryBuilder::default().category(3u8).position(4u16).description("Kommunikation").approved(1200.0).income_report(600.0).income_total(600.0).remark("Internet + Telefon").build()?)
                // Kategorie 4: Projektaktivitäten (3 Positionen) — Summe D=10500, E=F=10500
                .add_position(PositionEntryBuilder::default().category(4u8).position(1u16).description("Laptops").approved(6000.0).income_report(6000.0).income_total(6000.0).remark("4 Stück angeschafft").build()?)
                .add_position(PositionEntryBuilder::default().category(4u8).position(2u16).description("Beamer").approved(1500.0).income_report(1500.0).income_total(1500.0).remark("1 Stück").build()?)
                .add_position(PositionEntryBuilder::default().category(4u8).position(3u16).description("Mobiliar").approved(3000.0).income_report(3000.0).income_total(3000.0).remark("Tische + Stühle").build()?)
                // Kategorie 5: Projektverwaltung (2 Positionen) — Summe D=6000, E=F=3000
                .add_position(PositionEntryBuilder::default().category(5u8).position(1u16).description("Lehrerfortbildung").approved(4000.0).income_report(2000.0).income_total(2000.0).remark("2 Seminare").build()?)
                .add_position(PositionEntryBuilder::default().category(5u8).position(2u16).description("Supervision").approved(2000.0).income_report(1000.0).income_total(1000.0).remark("Quartalstreffen").build()?)
                // Kategorien 6-8: Header-Eingabe (position = 0)
                .add_position(PositionEntryBuilder::default().category(6u8).position(0u16).approved(8000.0).income_report(4000.0).income_total(4000.0).remark("Verwaltungspauschale").build()?)
                .add_position(PositionEntryBuilder::default().category(7u8).position(0u16).approved(3000.0).income_report(1500.0).income_total(1500.0).remark("Broschüren + Website").build()?)
                .add_position(PositionEntryBuilder::default().category(8u8).position(0u16).approved(1200.0).income_report(750.0).income_total(750.0).remark("Bankgebühren").build()?)
                // body_positions
                .body_positions(
                    [(1u8, 5u16), (2, 3), (3, 4), (4, 3), (5, 2), (6, 0), (7, 0), (8, 0)]
                        .into_iter()
                        .collect::<std::collections::HashMap<_, _>>()
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
                .locked(true)
                .hide_columns_qv(true)
                .hide_language_sheet(true)
                .build()?
        )
        .build()?;

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
