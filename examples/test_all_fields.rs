//! Test: Finanzbericht mit ALLEN Feldern gefüllt (E = F)
//!
//! Generiert einen vollständigen Finanzbericht, bei dem:
//! - Alle Header-Felder gesetzt sind
//! - Alle 5 Einnahmen-Zeilen gefüllt sind
//! - Alle 8 Kostenkategorien (5 mit Positionen, 3 Header-Eingabe) gefüllt sind
//! - Alle 36 Belege im Right Panel gefüllt sind
//! - Alle Footer-Felder (Bank, Kasse, Sonstiges) gesetzt sind
//! - Spalte E (Berichtsperiode) = Spalte F (Kumulativ) — wie im 1. Berichtszeitraum

use kmw_fb_rust::{PanelEntry, PositionEntry, ReportConfig, TableEntry};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle vollständigen Finanzbericht (E = F)...\n");

    let config = ReportConfig {
        language: "deutsch".to_string(),
        currency: "EUR".to_string(),
        project_number: Some("PROJ-2024-001".to_string()),
        project_title: Some("Bildungsprojekt Ostafrika".to_string()),
        project_start: Some("01.01.2024".to_string()),
        project_end: Some("31.12.2026".to_string()),
        report_start: Some("01.01.2024".to_string()),
        report_end: Some("30.06.2024".to_string()),

        // =====================================================================
        // Einnahmen-Tabelle (Rows 15-19) — E = F
        // =====================================================================
        table: vec![
            TableEntry {
                index: 0,
                approved_budget: Some(80000.0),
                income_report: Some(50000.0),
                income_total: Some(50000.0),
                reason: Some("1. Rate erhalten".to_string()),
            },
            TableEntry {
                index: 1,
                approved_budget: Some(15000.0),
                income_report: Some(7500.0),
                income_total: Some(7500.0),
                reason: Some("Eigenanteil eingebracht".to_string()),
            },
            TableEntry {
                index: 2,
                approved_budget: Some(10000.0),
                income_report: Some(5000.0),
                income_total: Some(5000.0),
                reason: Some("Stiftung XY".to_string()),
            },
            TableEntry {
                index: 3,
                approved_budget: Some(2500.0),
                income_report: Some(1500.0),
                income_total: Some(1500.0),
                reason: Some("Spenden".to_string()),
            },
            TableEntry {
                index: 4,
                approved_budget: Some(500.0),
                income_report: Some(400.0),
                income_total: Some(400.0),
                reason: Some("Tagesgeldkonto".to_string()),
            },
        ],

        // =====================================================================
        // Linkes Panel (Belege 1-18)
        // =====================================================================
        left_panel: vec![
            PanelEntry { index: 0, date: Some("15.01.2024".to_string()), amount_euro: Some(9000.0), amount_local: Some(9000.0) },
            PanelEntry { index: 1, date: Some("15.01.2024".to_string()), amount_euro: Some(6000.0), amount_local: Some(6000.0) },
            PanelEntry { index: 2, date: Some("31.01.2024".to_string()), amount_euro: Some(4000.0), amount_local: Some(4000.0) },
            PanelEntry { index: 3, date: Some("31.01.2024".to_string()), amount_euro: Some(3000.0), amount_local: Some(3000.0) },
            PanelEntry { index: 4, date: Some("15.02.2024".to_string()), amount_euro: Some(2000.0), amount_local: Some(2000.0) },
            PanelEntry { index: 5, date: Some("28.02.2024".to_string()), amount_euro: Some(2500.0), amount_local: Some(2500.0) },
            PanelEntry { index: 6, date: Some("01.03.2024".to_string()), amount_euro: Some(1500.0), amount_local: Some(1500.0) },
            PanelEntry { index: 7, date: Some("15.03.2024".to_string()), amount_euro: Some(750.0), amount_local: Some(750.0) },
            PanelEntry { index: 8, date: Some("31.03.2024".to_string()), amount_euro: Some(1000.0), amount_local: Some(1000.0) },
            PanelEntry { index: 9, date: Some("01.04.2024".to_string()), amount_euro: Some(2500.0), amount_local: Some(2500.0) },
            PanelEntry { index: 10, date: Some("15.04.2024".to_string()), amount_euro: Some(1800.0), amount_local: Some(1800.0) },
            PanelEntry { index: 11, date: Some("30.04.2024".to_string()), amount_euro: Some(600.0), amount_local: Some(600.0) },
            PanelEntry { index: 12, date: Some("01.05.2024".to_string()), amount_euro: Some(6000.0), amount_local: Some(6000.0) },
            PanelEntry { index: 13, date: Some("15.05.2024".to_string()), amount_euro: Some(1500.0), amount_local: Some(1500.0) },
            PanelEntry { index: 14, date: Some("31.05.2024".to_string()), amount_euro: Some(3000.0), amount_local: Some(3000.0) },
            PanelEntry { index: 15, date: Some("01.06.2024".to_string()), amount_euro: Some(2000.0), amount_local: Some(2000.0) },
            PanelEntry { index: 16, date: Some("15.06.2024".to_string()), amount_euro: Some(1000.0), amount_local: Some(1000.0) },
            PanelEntry { index: 17, date: Some("30.06.2024".to_string()), amount_euro: Some(4000.0), amount_local: Some(4000.0) },
        ],

        // =====================================================================
        // Rechtes Panel (Belege 19-36)
        // =====================================================================
        right_panel: vec![
            PanelEntry { index: 0, date: Some("05.01.2024".to_string()), amount_euro: Some(1500.0), amount_local: Some(1500.0) },
            PanelEntry { index: 1, date: Some("10.01.2024".to_string()), amount_euro: Some(750.0), amount_local: Some(750.0) },
            PanelEntry { index: 2, date: Some("20.01.2024".to_string()), amount_euro: Some(500.0), amount_local: Some(500.0) },
            PanelEntry { index: 3, date: Some("05.02.2024".to_string()), amount_euro: Some(2000.0), amount_local: Some(2000.0) },
            PanelEntry { index: 4, date: Some("15.02.2024".to_string()), amount_euro: Some(1200.0), amount_local: Some(1200.0) },
            PanelEntry { index: 5, date: Some("28.02.2024".to_string()), amount_euro: Some(800.0), amount_local: Some(800.0) },
            PanelEntry { index: 6, date: Some("10.03.2024".to_string()), amount_euro: Some(950.0), amount_local: Some(950.0) },
            PanelEntry { index: 7, date: Some("20.03.2024".to_string()), amount_euro: Some(1100.0), amount_local: Some(1100.0) },
            PanelEntry { index: 8, date: Some("01.04.2024".to_string()), amount_euro: Some(1800.0), amount_local: Some(1800.0) },
            PanelEntry { index: 9, date: Some("15.04.2024".to_string()), amount_euro: Some(650.0), amount_local: Some(650.0) },
            PanelEntry { index: 10, date: Some("30.04.2024".to_string()), amount_euro: Some(2200.0), amount_local: Some(2200.0) },
            PanelEntry { index: 11, date: Some("10.05.2024".to_string()), amount_euro: Some(1350.0), amount_local: Some(1350.0) },
            PanelEntry { index: 12, date: Some("20.05.2024".to_string()), amount_euro: Some(900.0), amount_local: Some(900.0) },
            PanelEntry { index: 13, date: Some("01.06.2024".to_string()), amount_euro: Some(1600.0), amount_local: Some(1600.0) },
            PanelEntry { index: 14, date: Some("15.06.2024".to_string()), amount_euro: Some(2100.0), amount_local: Some(2100.0) },
            PanelEntry { index: 15, date: Some("25.06.2024".to_string()), amount_euro: Some(750.0), amount_local: Some(750.0) },
            PanelEntry { index: 16, date: Some("28.06.2024".to_string()), amount_euro: Some(1400.0), amount_local: Some(1400.0) },
            PanelEntry { index: 17, date: Some("30.06.2024".to_string()), amount_euro: Some(850.0), amount_local: Some(850.0) },
        ],

        // =====================================================================
        // Kostenpositionen — E = F überall
        // =====================================================================
        positions: vec![
            // Kategorie 1: Personalkosten (5 Positionen) — Summe D=48000, E=F=24000
            PositionEntry { category: 1, position: 1, description: Some("Projektleitung".to_string()), approved: Some(18000.0), income_report: Some(9000.0), income_total: Some(9000.0), remark: Some("6 Monate".to_string()) },
            PositionEntry { category: 1, position: 2, description: Some("Buchhaltung".to_string()), approved: Some(12000.0), income_report: Some(6000.0), income_total: Some(6000.0), remark: Some("Teilzeit".to_string()) },
            PositionEntry { category: 1, position: 3, description: Some("Lokale Fachkräfte".to_string()), approved: Some(8000.0), income_report: Some(4000.0), income_total: Some(4000.0), remark: Some("2 Mitarbeiter".to_string()) },
            PositionEntry { category: 1, position: 4, description: Some("Sozialabgaben".to_string()), approved: Some(6000.0), income_report: Some(3000.0), income_total: Some(3000.0), remark: Some("AG-Anteile".to_string()) },
            PositionEntry { category: 1, position: 5, description: Some("Honorare Trainer".to_string()), approved: Some(4000.0), income_report: Some(2000.0), income_total: Some(2000.0), remark: Some("3 Workshops".to_string()) },
            // Kategorie 2: Reisekosten (3 Positionen) — Summe D=9500, E=F=4750
            PositionEntry { category: 2, position: 1, description: Some("Flüge International".to_string()), approved: Some(5000.0), income_report: Some(2500.0), income_total: Some(2500.0), remark: Some("2 Dienstreisen".to_string()) },
            PositionEntry { category: 2, position: 2, description: Some("Unterkunft/Verpflegung".to_string()), approved: Some(3000.0), income_report: Some(1500.0), income_total: Some(1500.0), remark: Some("Tagessätze".to_string()) },
            PositionEntry { category: 2, position: 3, description: Some("Lokale Transporte".to_string()), approved: Some(1500.0), income_report: Some(750.0), income_total: Some(750.0), remark: Some("Projektfahrten".to_string()) },
            // Kategorie 3: Sachkosten (4 Positionen) — Summe D=11800, E=F=5900
            PositionEntry { category: 3, position: 1, description: Some("Büromaterial".to_string()), approved: Some(2000.0), income_report: Some(1000.0), income_total: Some(1000.0), remark: Some("Verbrauchsmaterial".to_string()) },
            PositionEntry { category: 3, position: 2, description: Some("Lehrmaterialien".to_string()), approved: Some(5000.0), income_report: Some(2500.0), income_total: Some(2500.0), remark: Some("Schulbücher + Hefte".to_string()) },
            PositionEntry { category: 3, position: 3, description: Some("Miete Schulungsraum".to_string()), approved: Some(3600.0), income_report: Some(1800.0), income_total: Some(1800.0), remark: Some("6 Monate".to_string()) },
            PositionEntry { category: 3, position: 4, description: Some("Kommunikation".to_string()), approved: Some(1200.0), income_report: Some(600.0), income_total: Some(600.0), remark: Some("Internet + Telefon".to_string()) },
            // Kategorie 4: Investitionen (3 Positionen) — Summe D=10500, E=F=10500
            PositionEntry { category: 4, position: 1, description: Some("Laptops".to_string()), approved: Some(6000.0), income_report: Some(6000.0), income_total: Some(6000.0), remark: Some("4 Stück angeschafft".to_string()) },
            PositionEntry { category: 4, position: 2, description: Some("Beamer".to_string()), approved: Some(1500.0), income_report: Some(1500.0), income_total: Some(1500.0), remark: Some("1 Stück".to_string()) },
            PositionEntry { category: 4, position: 3, description: Some("Mobiliar".to_string()), approved: Some(3000.0), income_report: Some(3000.0), income_total: Some(3000.0), remark: Some("Tische + Stühle".to_string()) },
            // Kategorie 5: Fortbildung (2 Positionen) — Summe D=6000, E=F=3000
            PositionEntry { category: 5, position: 1, description: Some("Lehrerfortbildung".to_string()), approved: Some(4000.0), income_report: Some(2000.0), income_total: Some(2000.0), remark: Some("2 Seminare".to_string()) },
            PositionEntry { category: 5, position: 2, description: Some("Supervision".to_string()), approved: Some(2000.0), income_report: Some(1000.0), income_total: Some(1000.0), remark: Some("Quartalstreffen".to_string()) },
            // Kategorien 6-8: Header-Eingabe (position = 0)
            PositionEntry { category: 6, position: 0, description: None, approved: Some(8000.0), income_report: Some(4000.0), income_total: Some(4000.0), remark: Some("Verwaltungspauschale".to_string()) },
            PositionEntry { category: 7, position: 0, description: None, approved: Some(3000.0), income_report: Some(1500.0), income_total: Some(1500.0), remark: Some("Broschüren + Website".to_string()) },
            PositionEntry { category: 8, position: 0, description: None, approved: Some(1200.0), income_report: Some(750.0), income_total: Some(750.0), remark: Some("Bankgebühren".to_string()) },
        ],

        body_positions: [(1u8, 5u16), (2, 3), (3, 4), (4, 3), (5, 2), (6, 0), (7, 0), (8, 0)]
            .into_iter()
            .collect(),

        footer_bank: Some(8500.0),
        footer_kasse: Some(1200.0),
        footer_sonstiges: Some(300.0),

        locked: true,
        hide_columns_qv: true,
        hide_language_sheet: true,

        ..ReportConfig::default()
    };

    let output_path = "tests/output/test_all_fields.xlsx";
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
