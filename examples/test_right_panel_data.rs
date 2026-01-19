//! Test: Right Panel mit vollständigen Daten (Datum + Beträge)
//!
//! Erstellt Testdateien mit:
//! - L14-L31: Datumswerte
//! - M14-M31: Beträge in Euro
//! - N14-N31: Beträge in Lokalwährung
//! - K14-K31, R14-R31: Nummern

use kmw_fb_rust::v2::report::formats::ReportStyles;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::values::ReportValues;
use kmw_fb_rust::v2::report::writer::write_report_v2;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔══════════════════════════════════════════════════════════════╗");
    println!("║        Test: Right Panel mit vollständigen Daten             ║");
    println!("╚══════════════════════════════════════════════════════════════╝");
    println!();

    let mut workbook = Workbook::new();

    // 1. Add Target Sheet
    let sheet_name = "Finanzbericht";
    let _ = workbook.add_worksheet().set_name(sheet_name)?;

    // 2. Build Translation Sheet
    build_trans_sheet(&mut workbook)?;

    // 3. Get Target Sheet back
    let ws = workbook.worksheet_from_name(sheet_name)?;

    // 4. Set column format to unlocked
    let unlocked = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked();

    for col in 0..1000u16 {
        ws.set_column_format(col, &unlocked)?;
    }

    // 5. Setup sheet
    setup_sheet(ws)?;

    // 6. Prepare Styles
    let styles = ReportStyles::new();

    // 7. Prepare Values mit vollständigen Daten
    let mut values = ReportValues::new()
        .with_language("deutsch")
        .with_currency("EUR")
        .with_project_number("TEST-RIGHT-PANEL")
        .with_project_title("Test Right Panel mit Daten")
        .with_project_start("2024-01-01")
        .with_project_end("2024-12-31")
        .with_report_start("2024-01-01")
        .with_report_end("2024-03-31")
        .with_exchange_rate(1.1);

    // Tabellendaten setzen (D15-F19)
    use kmw_fb_rust::v2::report::cells::{RightPanelInputCell, TableInputCell};

    println!("📝 Setze Tabellendaten (D15-F19)...");
    for i in 0..5u8 {
        values
            .set(TableInputCell::ApprovedBudget(i), 1000.0 + i as f64 * 500.0)
            .set(
                TableInputCell::IncomeReportPeriod(i),
                100.0 + i as f64 * 200.0,
            )
            .set(TableInputCell::IncomeTotal(i), 500.0 + i as f64 * 300.0);
    }

    // Right Panel Daten setzen (K14-U31)
    println!("📝 Setze Right Panel Daten (K14-K31, L14-L31, M14-M31, N14-N31)...");
    for i in 0..18u8 {
        let num = i + 1;
        let amount_eur = 100.0 + i as f64 * 250.0;
        let amount_local = 110.0 + i as f64 * 275.0;

        // Datumswerte (verschiedene Daten im März/April/Mai 2024)
        let day = (i % 28) + 1;
        let month = 3 + (i / 10);
        let date_str = format!("2024-{:02}-{:02}", month, day);

        values
            .set(RightPanelInputCell::LeftNumber(i), num.to_string())
            .set(RightPanelInputCell::LeftDate(i), date_str.clone())
            .set(RightPanelInputCell::LeftAmount1(i), amount_eur)
            .set(RightPanelInputCell::LeftAmount2(i), amount_local)
            .set(RightPanelInputCell::RightNumber(i), (num + 100).to_string())
            .set(RightPanelInputCell::RightDate(i), date_str)
            .set(RightPanelInputCell::RightAmount1(i), amount_eur * 1.5)
            .set(RightPanelInputCell::RightAmount2(i), amount_local * 1.5);
    }

    println!("✓ Alle Daten gesetzt");
    println!();

    // 8. Write Report
    println!("📄 Schreibe Excel-Report...");
    let suffix = "_de";
    write_report_v2(ws, &styles, suffix, &values)?;

    // 9. Protect worksheet
    ws.protect();

    // 10. Save to file
    let path = "test_right_panel_complete.xlsx";
    workbook.save(path)?;
    println!("✓ Datei gespeichert: {}", path);
    println!();

    println!("═══════════════════════════════════════════════════════════════");
    println!("                   DATEN ÜBERSICHT");
    println!("═══════════════════════════════════════════════════════════════");
    println!();

    println!("📊 Tabellen-Daten (D15-F19):");
    println!("   Budget (D):  1000, 1500, 2000, 2500, 3000");
    println!("   Einnahmen (E): 100, 300, 500, 700, 900");
    println!("   Gesamt (F):  500, 800, 1100, 1400, 1700");
    println!();

    println!("📊 Right Panel Daten (L14-L31 Daten, M14-M31 Euro, N14-N31 Lokal):");
    println!("   Insgesamt 18 Zeilen mit:");
    println!("     - Datumswerte (variierend)");
    println!("     - Euro-Beträge (100-5450)");
    println!("     - Lokale Beträge (110-5995)");
    println!("   Linke Seite (K-N):  Nummern 1-18, Daten wie oben");
    println!("   Rechte Seite (R-U): Nummern 101-118, 1.5x die linken Beträge");
    println!();

    println!("✅ Test abgeschlossen!");

    Ok(())
}
