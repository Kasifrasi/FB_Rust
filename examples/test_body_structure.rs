//! Test: Body-Struktur visuell prüfen
//!
//! Generiert eine Excel-Datei mit dem dynamischen Body-Bereich.
//! Öffne die Datei in Excel um das Layout zu prüfen.

use kmw_fb_rust::v2::lang::build_sheet as build_sprachversionen;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::ApiKey;
use kmw_fb_rust::v2::report::{write_report_v2_with_body, BodyConfig, ReportStyles, ReportValues};
use kmw_fb_rust::v2::report::{PositionField, SingleRowField};
use rust_xlsxwriter::Workbook;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle Test-Datei für Body-Struktur...");

    let mut workbook = Workbook::new();
    let styles = ReportStyles::new();

    // Sprachversionen-Sheet hinzufügen (für VLOOKUP-Formeln)
    build_sprachversionen(&mut workbook)?;

    // Worksheet erstellen und einrichten
    let ws = workbook.add_worksheet();
    ws.set_name("Finanzbericht")?;
    setup_sheet(ws)?;

    // Test-Werte setzen
    let mut values = ReportValues::new();
    values.set(ApiKey::Language, "deutsch");
    values.set(ApiKey::Currency, "EUR");
    values.set(ApiKey::ProjectNumber, "TEST-001");
    values.set(ApiKey::ProjectTitle, "Test Body Structure");
    values.set(ApiKey::ProjectStart, "01.01.2024");
    values.set(ApiKey::ProjectEnd, "31.12.2024");
    values.set(ApiKey::ReportStart, "01.01.2024");
    values.set(ApiKey::ReportEnd, "30.06.2024");

    // Einnahmen-Tabelle (statisch)
    for i in 0..5u8 {
        values.set(ApiKey::ApprovedBudget(i), 10000.0 * (i + 1) as f64);
        values.set(ApiKey::IncomeReportPeriod(i), 5000.0 * (i + 1) as f64);
        values.set(ApiKey::IncomeTotal(i), 5000.0 * (i + 1) as f64);
    }

    // Body-Konfiguration mit weniger Positionen für bessere Übersicht
    let body_config = BodyConfig::new()
        .with_positions(1, 5) // Kategorie 1: 5 Positionen
        .with_positions(2, 3) // Kategorie 2: 3 Positionen
        .with_positions(3, 4) // Kategorie 3: 4 Positionen
        .with_positions(4, 3) // Kategorie 4: 3 Positionen
        .with_positions(5, 2); // Kategorie 5: 2 Positionen

    // =========================================================================
    // Kostenpositionen (dynamisch via API)
    // =========================================================================

    // Kategorie 1: Personalkosten (5 Positionen)
    values.set_position_row(1, 1, "Projektleitung", 15000.0, 7500.0, 7500.0, "");
    values.set_position_row(
        1,
        2,
        "Wissenschaftliche Mitarbeiter",
        25000.0,
        12000.0,
        12000.0,
        "",
    );
    values.set_position_row(1, 3, "Studentische Hilfskräfte", 5000.0, 2500.0, 2500.0, "");
    values.set_position_row(
        1,
        4,
        "Externe Berater",
        8000.0,
        4000.0,
        4000.0,
        "Beratung IT",
    );
    values.set_position_row(1, 5, "Sonstige Personalkosten", 2000.0, 1000.0, 1000.0, "");

    // Kategorie 2: Sachkosten (3 Positionen)
    values.set_position_row(2, 1, "Büromaterial", 1500.0, 800.0, 800.0, "");
    values.set_position_row(
        2,
        2,
        "IT-Ausstattung",
        3000.0,
        2000.0,
        2000.0,
        "Laptop + Monitor",
    );
    values.set_position_row(2, 3, "Software-Lizenzen", 2000.0, 1500.0, 1500.0, "");

    // Kategorie 3: Reisekosten (4 Positionen)
    values.set_position_row(3, 1, "Inlandsreisen", 2000.0, 1200.0, 1200.0, "");
    values.set_position_row(
        3,
        2,
        "Auslandsreisen",
        5000.0,
        2500.0,
        2500.0,
        "Konferenz Berlin",
    );
    values.set_position_row(3, 3, "Tagegeld", 1000.0, 500.0, 500.0, "");
    values.set_position_row(3, 4, "Fahrtkosten", 800.0, 400.0, 400.0, "");

    // Kategorie 4: Investitionen (3 Positionen)
    values.set_position_row(4, 1, "Laborgeräte", 10000.0, 5000.0, 5000.0, "");
    values.set_position_row(4, 2, "Messgeräte", 8000.0, 4000.0, 4000.0, "Anschaffung Q2");
    values.set_position_row(4, 3, "Sonstige Investitionen", 2000.0, 1000.0, 1000.0, "");

    // Kategorie 5: Sonstige Kosten (2 Positionen)
    values.set_position_row(
        5,
        1,
        "Publikationskosten",
        3000.0,
        1500.0,
        1500.0,
        "Open Access",
    );
    values.set_position_row(5, 2, "Allgemeine Verwaltung", 1500.0, 750.0, 750.0, "");

    // Einzelne Position mit individuellem Feld setzen (zur Demonstration)
    values.set_position(1, 1, PositionField::Remark, "Senior Developer");

    // =========================================================================
    // Single-Row Kategorien (6, 7, 8) - nur eine Zeile pro Kategorie
    // =========================================================================

    // Kategorie 6: Sonstige direkte Kosten
    values.set_single_row_values(6, 4000.0, 2000.0, 2000.0, "Diverse");

    // Kategorie 7: Indirekte Kosten
    values.set_single_row_values(7, 6000.0, 3000.0, 3000.0, "Overhead");

    // Kategorie 8: Eigenmittel/Drittmittel (Einzelfeld-Demonstration)
    values.set_single_row(8, SingleRowField::Approved, 2500.0);
    values.set_single_row(8, SingleRowField::IncomeReport, 1250.0);
    values.set_single_row(8, SingleRowField::IncomeTotal, 1250.0);
    values.set_single_row(8, SingleRowField::Remark, "Eigenbeitrag");

    println!("Positions-Werte gesetzt für {} Multi-Row Kategorien", 5);
    println!("Single-Row-Werte gesetzt für {} Kategorien (6, 7, 8)", 3);

    // Report schreiben
    let result = write_report_v2_with_body(ws, &styles, "v2-body-test", &values, &body_config)?;

    println!("Body-Layout:");
    println!(
        "  Total Row: {} (Excel Row {})",
        result.total_row,
        result.total_row + 1
    );
    println!(
        "  Last Row: {} (Excel Row {})",
        result.last_row,
        result.last_row + 1
    );
    println!("  Kategorien:");
    for cat in &result.layout.categories {
        if let Some(header) = cat.header_row {
            let positions = cat.positions.as_ref().unwrap();
            let footer = cat.footer_row.unwrap();
            println!(
                "    Kategorie {}: Header={}, Pos={}-{} ({}), Footer={}",
                cat.meta.num,
                header + 1,
                positions.start_row + 1,
                positions.end_row + 1,
                positions.count,
                footer + 1
            );
        } else if let Some(single) = cat.single_row {
            println!("    Kategorie {}: Single-Row={}", cat.meta.num, single + 1);
        }
    }

    // Speichern
    let path = "body_structure_test.xlsx";
    workbook.save(path)?;
    println!("\nDatei gespeichert: {}", path);
    println!("Öffne die Datei in Excel um das Layout zu prüfen.");

    Ok(())
}
