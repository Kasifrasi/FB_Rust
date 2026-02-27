//! Test: Finanzbericht mit ALLEN Feldern gefüllt (E = F)
//!
//! Generiert einen vollständigen Finanzbericht, bei dem:
//! - Alle Header-Felder gesetzt sind
//! - Alle 5 Einnahmen-Zeilen gefüllt sind
//! - Alle 8 Kostenkategorien (5 mit Positionen, 3 Header-Eingabe) gefüllt sind
//! - Alle 36 Belege im Right Panel gefüllt sind
//! - Alle Footer-Felder (Bank, Kasse, Sonstiges) gesetzt sind
//! - Spalte E (Berichtsperiode) = Spalte F (Kumulativ) — wie im 1. Berichtszeitraum

use kmw_fb_rust::report::CellValue;
use kmw_fb_rust::{
    create_protected_report, ApiKey, BodyConfig, ReportOptions, ReportStyles, ReportValues,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle vollständigen Finanzbericht (E = F)...\n");

    let styles = ReportStyles::new();
    let mut values = ReportValues::new();

    // =========================================================================
    // Header (Rows 1-9)
    // =========================================================================

    values.set(ApiKey::Language, "deutsch");
    values.set(ApiKey::Currency, "EUR");
    values.set(ApiKey::ProjectNumber, "PROJ-2024-001");
    values.set(ApiKey::ProjectTitle, "Bildungsprojekt Ostafrika");
    values.set(ApiKey::ProjectStart, "01.01.2024");
    values.set(ApiKey::ProjectEnd, "31.12.2026");
    values.set(ApiKey::ReportStart, "01.01.2024");
    values.set(ApiKey::ReportEnd, "30.06.2024");

    // =========================================================================
    // Einnahmen-Tabelle (Rows 15-19) — E = F
    // =========================================================================

    // Zeile 1: KMW-Zuschuss
    values.set(ApiKey::ApprovedBudget(0), 80000.0);
    values.set(ApiKey::IncomeReportPeriod(0), 50000.0);
    values.set(ApiKey::IncomeTotal(0), 50000.0);
    values.set(ApiKey::IncomeReason(0), "1. Rate erhalten");

    // Zeile 2: Eigenmittel
    values.set(ApiKey::ApprovedBudget(1), 15000.0);
    values.set(ApiKey::IncomeReportPeriod(1), 7500.0);
    values.set(ApiKey::IncomeTotal(1), 7500.0);
    values.set(ApiKey::IncomeReason(1), "Eigenanteil eingebracht");

    // Zeile 3: Drittmittel
    values.set(ApiKey::ApprovedBudget(2), 10000.0);
    values.set(ApiKey::IncomeReportPeriod(2), 5000.0);
    values.set(ApiKey::IncomeTotal(2), 5000.0);
    values.set(ApiKey::IncomeReason(2), "Stiftung XY");

    // Zeile 4: Sonstige Einnahmen
    values.set(ApiKey::ApprovedBudget(3), 2500.0);
    values.set(ApiKey::IncomeReportPeriod(3), 1500.0);
    values.set(ApiKey::IncomeTotal(3), 1500.0);
    values.set(ApiKey::IncomeReason(3), "Spenden");

    // Zeile 5: Zinserträge
    values.set(ApiKey::ApprovedBudget(4), 500.0);
    values.set(ApiKey::IncomeReportPeriod(4), 400.0);
    values.set(ApiKey::IncomeTotal(4), 400.0);
    values.set(ApiKey::IncomeReason(4), "Tagesgeldkonto");

    // =========================================================================
    // Body-Konfiguration: 8 Kategorien
    // =========================================================================

    let body_config = BodyConfig::new()
        .with_positions(1, 5) // Personalkosten: 5 Positionen
        .with_positions(2, 3) // Reisekosten: 3 Positionen
        .with_positions(3, 4) // Sachkosten: 4 Positionen
        .with_positions(4, 3) // Investitionen: 3 Positionen
        .with_positions(5, 2) // Fortbildung: 2 Positionen
        .with_positions(6, 0) // Verwaltung: Header-Eingabe
        .with_positions(7, 0) // Öffentlichkeitsarbeit: Header-Eingabe
        .with_positions(8, 0); // Sonstiges: Header-Eingabe

    // =========================================================================
    // Kostenpositionen — E = F überall
    // =========================================================================

    // --- Kategorie 1: Personalkosten (5 Positionen) ---
    values.set_position_row(1, 1, "Projektleitung", 18000.0, 9000.0, 9000.0, "6 Monate");
    values.set_position_row(1, 2, "Buchhaltung", 12000.0, 6000.0, 6000.0, "Teilzeit");
    values.set_position_row(
        1,
        3,
        "Lokale Fachkräfte",
        8000.0,
        4000.0,
        4000.0,
        "2 Mitarbeiter",
    );
    values.set_position_row(
        1,
        4,
        "Sozialabgaben",
        6000.0,
        3000.0,
        3000.0,
        "AG-Anteile",
    );
    values.set_position_row(1, 5, "Honorare Trainer", 4000.0, 2000.0, 2000.0, "3 Workshops");
    // Summe Kat 1: D=48000, E=F=24000

    // --- Kategorie 2: Reisekosten (3 Positionen) ---
    values.set_position_row(
        2,
        1,
        "Flüge International",
        5000.0,
        2500.0,
        2500.0,
        "2 Dienstreisen",
    );
    values.set_position_row(
        2,
        2,
        "Unterkunft/Verpflegung",
        3000.0,
        1500.0,
        1500.0,
        "Tagessätze",
    );
    values.set_position_row(
        2,
        3,
        "Lokale Transporte",
        1500.0,
        750.0,
        750.0,
        "Projektfahrten",
    );
    // Summe Kat 2: D=9500, E=F=4750

    // --- Kategorie 3: Sachkosten (4 Positionen) ---
    values.set_position_row(
        3,
        1,
        "Büromaterial",
        2000.0,
        1000.0,
        1000.0,
        "Verbrauchsmaterial",
    );
    values.set_position_row(
        3,
        2,
        "Lehrmaterialien",
        5000.0,
        2500.0,
        2500.0,
        "Schulbücher + Hefte",
    );
    values.set_position_row(3, 3, "Miete Schulungsraum", 3600.0, 1800.0, 1800.0, "6 Monate");
    values.set_position_row(
        3,
        4,
        "Kommunikation",
        1200.0,
        600.0,
        600.0,
        "Internet + Telefon",
    );
    // Summe Kat 3: D=11800, E=F=5900

    // --- Kategorie 4: Investitionen (3 Positionen) ---
    values.set_position_row(4, 1, "Laptops", 6000.0, 6000.0, 6000.0, "4 Stück angeschafft");
    values.set_position_row(4, 2, "Beamer", 1500.0, 1500.0, 1500.0, "1 Stück");
    values.set_position_row(4, 3, "Mobiliar", 3000.0, 3000.0, 3000.0, "Tische + Stühle");
    // Summe Kat 4: D=10500, E=F=10500

    // --- Kategorie 5: Fortbildung (2 Positionen) ---
    values.set_position_row(
        5,
        1,
        "Lehrerfortbildung",
        4000.0,
        2000.0,
        2000.0,
        "2 Seminare",
    );
    values.set_position_row(
        5,
        2,
        "Supervision",
        2000.0,
        1000.0,
        1000.0,
        "Quartalstreffen",
    );
    // Summe Kat 5: D=6000, E=F=3000

    // --- Kategorie 6: Verwaltung (Header-Eingabe) ---
    values.set_header_input(6, 8000.0, 4000.0, 4000.0, "Verwaltungspauschale");
    // Summe Kat 6: D=8000, E=F=4000

    // --- Kategorie 7: Öffentlichkeitsarbeit (Header-Eingabe) ---
    values.set_header_input(7, 3000.0, 1500.0, 1500.0, "Broschüren + Website");
    // Summe Kat 7: D=3000, E=F=1500

    // --- Kategorie 8: Sonstiges (Header-Eingabe) ---
    values.set_header_input(8, 1200.0, 750.0, 750.0, "Bankgebühren");
    // Summe Kat 8: D=1200, E=F=750

    // Gesamtausgaben: D=98000, E=F=54400

    // =========================================================================
    // Right Panel: Belege (36 Einträge)
    // =========================================================================

    // --- Linkes Panel (Beleg 1-18) ---
    let left_entries: &[(&str, f64, f64)] = &[
        ("15.01.2024", 9000.0, 9000.0),
        ("15.01.2024", 6000.0, 6000.0),
        ("31.01.2024", 4000.0, 4000.0),
        ("31.01.2024", 3000.0, 3000.0),
        ("15.02.2024", 2000.0, 2000.0),
        ("28.02.2024", 2500.0, 2500.0),
        ("01.03.2024", 1500.0, 1500.0),
        ("15.03.2024", 750.0, 750.0),
        ("31.03.2024", 1000.0, 1000.0),
        ("01.04.2024", 2500.0, 2500.0),
        ("15.04.2024", 1800.0, 1800.0),
        ("30.04.2024", 600.0, 600.0),
        ("01.05.2024", 6000.0, 6000.0),
        ("15.05.2024", 1500.0, 1500.0),
        ("31.05.2024", 3000.0, 3000.0),
        ("01.06.2024", 2000.0, 2000.0),
        ("15.06.2024", 1000.0, 1000.0),
        ("30.06.2024", 4000.0, 4000.0),
    ];

    for (i, (date, euro, local)) in left_entries.iter().enumerate() {
        let idx = i as u8;
        values.set(ApiKey::LeftDate(idx), CellValue::Date(date.to_string()));
        values.set(ApiKey::LeftAmountEuro(idx), *euro);
        values.set(ApiKey::LeftAmountLocal(idx), *local);
    }

    // --- Rechtes Panel (Beleg 19-36) ---
    let right_entries: &[(&str, f64, f64)] = &[
        ("05.01.2024", 1500.0, 1500.0),
        ("10.01.2024", 750.0, 750.0),
        ("20.01.2024", 500.0, 500.0),
        ("05.02.2024", 2000.0, 2000.0),
        ("15.02.2024", 1200.0, 1200.0),
        ("28.02.2024", 800.0, 800.0),
        ("10.03.2024", 950.0, 950.0),
        ("20.03.2024", 1100.0, 1100.0),
        ("01.04.2024", 1800.0, 1800.0),
        ("15.04.2024", 650.0, 650.0),
        ("30.04.2024", 2200.0, 2200.0),
        ("10.05.2024", 1350.0, 1350.0),
        ("20.05.2024", 900.0, 900.0),
        ("01.06.2024", 1600.0, 1600.0),
        ("15.06.2024", 2100.0, 2100.0),
        ("25.06.2024", 750.0, 750.0),
        ("28.06.2024", 1400.0, 1400.0),
        ("30.06.2024", 850.0, 850.0),
    ];

    for (i, (date, euro, local)) in right_entries.iter().enumerate() {
        let idx = i as u8;
        values.set(ApiKey::RightDate(idx), CellValue::Date(date.to_string()));
        values.set(ApiKey::RightAmountEuro(idx), *euro);
        values.set(ApiKey::RightAmountLocal(idx), *local);
    }

    // =========================================================================
    // Footer: Saldenabstimmung
    // =========================================================================

    values.set_footer_bank(8500.0);
    values.set_footer_kasse(1200.0);
    values.set_footer_sonstiges(300.0);

    // =========================================================================
    // Report generieren
    // =========================================================================

    let options = ReportOptions::with_default_protection()
        .with_hidden_columns_qv()
        .with_hidden_language_sheet();

    let output_path = "tests/output/test_all_fields.xlsx";
    create_protected_report(output_path, &styles, &values, &body_config, &options)?;

    println!("Finanzbericht erstellt: {}", output_path);
    println!();
    println!("Zusammenfassung:");
    println!("  Header:          8 Felder gesetzt");
    println!("  Einnahmen:       5 Zeilen x 4 Spalten (D, E, F, H)");
    println!("  Kostenpositionen: 17 Positionen + 3 Header-Eingaben");
    println!("  Belege:          36 Einträge (18 links + 18 rechts)");
    println!("  Footer:          Bank={}, Kasse={}, Sonstiges={}", 8500.0, 1200.0, 300.0);
    println!();
    println!("  Spalte E = Spalte F (1. Berichtszeitraum = Kumulativ)");

    Ok(())
}
