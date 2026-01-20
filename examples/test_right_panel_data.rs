//! Test: Right Panel mit vollständigen Daten (Datum + Beträge)
//!
//! NUR diese Zellen sind API-Eingabefelder:
//! - L14:N31 (Datum, Euro, Lokal - linke Seite)
//! - S14:U31 (Datum, Euro, Lokal - rechte Seite)
//!
//! K14:K31 und R14:R31 sind FORMELN und dürfen NICHT befüllt werden!

use kmw_fb_rust::v2::report::api::ApiKey;
use kmw_fb_rust::v2::report::formats::ReportStyles;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::values::ReportValues;
use kmw_fb_rust::v2::report::writer::write_report_v2;
use kmw_fb_rust::v2::Sprachversion::builder::build_sheet as build_trans_sheet;
use rust_xlsxwriter::{Format, Workbook};

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
    // API-Felder: E2:E3, D5, D6, E8:E9, G8:G9, D15:F19, H15:H19, L14:N31, S14:U31
    let mut values = ReportValues::new()
        .with_language("deutsch")
        .with_currency("EUR")
        .with_project_number("TEST-RIGHT-PANEL")
        .with_project_title("Test Right Panel mit Daten")
        .with_project_start("2024-01-01")
        .with_project_end("2024-12-31")
        .with_report_start("2024-01-01")
        .with_report_end("2024-03-31");

    // Tabellendaten setzen (D15-F19)
    println!("Setze Tabellendaten (D15-F19)...");
    for i in 0..5u8 {
        values
            .set(ApiKey::ApprovedBudget(i), 1000.0 + i as f64 * 500.0)
            .set(ApiKey::IncomeReportPeriod(i), 100.0 + i as f64 * 200.0)
            .set(ApiKey::IncomeTotal(i), 500.0 + i as f64 * 300.0);
    }

    // Right Panel Daten setzen - NUR L, M, N (links) und S, T, U (rechts)
    // K14:K31 und R14:R31 sind FORMELN, nicht befüllen!
    println!("Setze Right Panel Daten (L14-N31, S14-U31)...");
    println!("   HINWEIS: K14:K31 und R14:R31 sind Formeln - nicht befuellen!");
    for i in 0..18u8 {
        let amount_eur = 100.0 + i as f64 * 250.0;
        let amount_local = 110.0 + i as f64 * 275.0;

        // Datumswerte (verschiedene Daten im März/April/Mai 2024)
        let day = (i % 28) + 1;
        let month = 3 + (i / 10);
        let date_str = format!("2024-{:02}-{:02}", month, day);

        values
            // Linke Seite: L, M, N (Datum, Euro, Lokal)
            .set(ApiKey::LeftDate(i), date_str.clone())
            .set(ApiKey::LeftAmountEuro(i), amount_eur)
            .set(ApiKey::LeftAmountLocal(i), amount_local)
            // Rechte Seite: S, T, U (Datum, Euro, Lokal)
            .set(ApiKey::RightDate(i), date_str)
            .set(ApiKey::RightAmountEuro(i), amount_eur * 1.5)
            .set(ApiKey::RightAmountLocal(i), amount_local * 1.5);
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

    println!("📊 API-Felder (von außen befüllbar):");
    println!("   E2:E3     - Sprache, Währung");
    println!("   D5, D6    - Projektnummer, Projekttitel");
    println!("   E8:E9     - Projektstart, Berichtszeitraum Start");
    println!("   G8:G9     - Projektende, Berichtszeitraum Ende");
    println!("   D15:F19   - Budget, Einnahmen Berichtszeitraum, Einnahmen Gesamt");
    println!("   H15:H19   - Begründung");
    println!("   L14:N31   - Datum, Euro, Lokal (linke Seite)");
    println!("   S14:U31   - Datum, Euro, Lokal (rechte Seite)");
    println!();

    println!("⚠️  FORMELN (NICHT von API befüllbar):");
    println!("   J7, J8    - Wechselkurs Datum/Wert");
    println!("   G15:G19   - Prozent (IFERROR Division)");
    println!("   K14:K31   - Nummern links (Formeln!)");
    println!("   R14:R31   - Nummern rechts (Formeln!)");
    println!();

    println!("✅ Test abgeschlossen!");

    Ok(())
}
