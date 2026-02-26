//! Test: Mehrsprachige Reports mit Gruppierung
//!
//! Erstellt Reports in allen 5 Sprachen mit:
//! - Bauausgaben: 20 Positionen (letzte 15 gruppiert)
//! - Einmalige Ausgaben: 20 Positionen (letzte 15 gruppiert)
//! - Personalkosten: 30 Positionen (letzte 18 gruppiert)
//! - Projektaktivitäten: 30 Positionen (letzte 18 gruppiert)
//! - Projektverwaltung: 20 Positionen (letzte 15 gruppiert)
//! - Evaluierung: 0 Positionen (nur Header)
//! - Audit: 0 Positionen (nur Header)
//! - Reserve: 0 Positionen (nur Header)

use kmw_fb_rust::report::writer::setup_sheet;
use kmw_fb_rust::{
    build_language_sheet_with_visibility, write_report_with_options, BodyConfig, ReportOptions,
    ReportStyles, ReportValues, RowGrouping,
};
use rust_xlsxwriter::Workbook;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let languages = ["deutsch", "english", "francais", "espanol", "portugues"];

    for lang in &languages {
        create_report(lang)?;
    }

    println!("\n✅ Alle 5 Reports erstellt!");
    Ok(())
}

fn create_report(language: &str) -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle Report für: {}", language);

    // 1. Workbook erstellen
    let mut workbook = Workbook::new();

    // 2. Sprachversionen-Sheet (ausgeblendet)
    build_language_sheet_with_visibility(&mut workbook, true)?;

    // 3. Haupt-Worksheet
    let ws = workbook.add_worksheet().set_name("Finanzbericht")?;

    // 4. Sheet Setup (Spaltenbreiten, Zeilenhöhen, Gitternetzlinien)
    setup_sheet(ws)?;

    // 5. Styles
    let styles = ReportStyles::new();

    // 6. Values - nur Sprache, keine Dummy-Werte
    let values = ReportValues::new().with_language(language);

    // 7. Body Config:
    //    Kat 1 (Bauausgaben): 20 Positionen
    //    Kat 2 (Einmalige Ausgaben): 20 Positionen
    //    Kat 3 (Personalkosten): 30 Positionen
    //    Kat 4 (Projektaktivitäten): 30 Positionen
    //    Kat 5 (Projektverwaltung): 20 Positionen
    //    Kat 6 (Evaluierung): 0 (Header-Eingabe)
    //    Kat 7 (Audit): 0 (Header-Eingabe)
    //    Kat 8 (Reserve): 0 (Header-Eingabe)
    let body_config = BodyConfig::new()
        .with_positions(1, 20)
        .with_positions(2, 20)
        .with_positions(3, 30)
        .with_positions(4, 30)
        .with_positions(5, 20)
        .with_positions(6, 0)
        .with_positions(7, 0)
        .with_positions(8, 0);

    // 8. Berechne Zeilen für Gruppierung
    //    Body startet bei Zeile 26 (0-basiert)
    //
    //    Kategorie 1 (20 Positionen):
    //      Header: 26
    //      Positionen: 27-46 (20 Zeilen)
    //      Footer: 47
    //      -> Gruppiere Zeilen 32-46 (Position 6-20, also letzte 15)
    //
    //    Kategorie 2 (20 Positionen):
    //      Header: 48
    //      Positionen: 49-68 (20 Zeilen)
    //      Footer: 69
    //      -> Gruppiere Zeilen 54-68 (Position 6-20, also letzte 15)
    //
    //    Kategorie 3 (30 Positionen):
    //      Header: 70
    //      Positionen: 71-100 (30 Zeilen)
    //      Footer: 101
    //      -> Gruppiere Zeilen 83-100 (Position 13-30, also letzte 18)
    //
    //    Kategorie 4 (30 Positionen):
    //      Header: 102
    //      Positionen: 103-132 (30 Zeilen)
    //      Footer: 133
    //      -> Gruppiere Zeilen 115-132 (Position 13-30, also letzte 18)
    //
    //    Kategorie 5 (20 Positionen):
    //      Header: 134
    //      Positionen: 135-154 (20 Zeilen)
    //      Footer: 155
    //      -> Gruppiere Zeilen 140-154 (Position 6-20, also letzte 15)
    //
    //    Kategorien 6, 7, 8: je 1 Zeile (Header-Eingabe), keine Gruppierung

    let row_grouping = RowGrouping::new()
        // Kategorie 1: Zeilen 32-46 (letzte 15 von 20)
        .add_collapsed_group(32, 46)
        // Kategorie 2: Zeilen 54-68 (letzte 15 von 20)
        .add_collapsed_group(54, 68)
        // Kategorie 3: Zeilen 83-100 (letzte 18 von 30)
        .add_collapsed_group(83, 100)
        // Kategorie 4: Zeilen 115-132 (letzte 18 von 30)
        .add_collapsed_group(115, 132)
        // Kategorie 5: Zeilen 140-154 (letzte 15 von 20)
        .add_collapsed_group(140, 154);

    // 9. Report Options - KEIN Schutz, nur Gruppierung
    let options = ReportOptions::new()
        .with_row_grouping(row_grouping)
        .with_hidden_language_sheet();

    // 10. Report schreiben
    write_report_with_options(ws, &styles, "", &values, &body_config, &options)?;

    // 11. Datei speichern
    let filename = format!("test_grouped_{}.xlsx", language);
    workbook.save(&filename)?;
    println!("  ✓ {}", filename);

    Ok(())
}
