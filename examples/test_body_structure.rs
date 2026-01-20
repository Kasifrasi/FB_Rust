//! Test: Body-Struktur visuell prüfen
//!
//! Generiert eine Excel-Datei mit dem dynamischen Body-Bereich.
//! Öffne die Datei in Excel um das Layout zu prüfen.

use kmw_fb_rust::v2::lang::build_sheet as build_sprachversionen;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::ApiKey;
use kmw_fb_rust::v2::report::{write_report_v2_with_body, BodyConfig, ReportStyles, ReportValues};
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
