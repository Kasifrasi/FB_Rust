//! Test: Mehrsprachige Finanzberichte
//!
//! Generiert 5 Excel-Dateien in verschiedenen Sprachen:
//! - Deutsch (D_Finanzbericht.xlsx)
//! - Englisch (E_Financial_report.xlsx)
//! - Französisch (F_Rapport_financier.xlsx)
//! - Spanisch (S_Informe_financiero.xlsx)
//! - Portugiesisch (P_Relatorio_financeiro.xlsx)

use kmw_fb_rust::v2::common::{LANG_CONFIG, LANG_SUFFIXES};
use kmw_fb_rust::v2::lang::build_sheet as build_sprachversionen;
use kmw_fb_rust::v2::report::layout::setup_sheet;
use kmw_fb_rust::v2::report::ApiKey;
use kmw_fb_rust::v2::report::{write_report_v2_with_body, BodyConfig, ReportStyles, ReportValues};
use rust_xlsxwriter::{Format, Workbook};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle mehrsprachige Finanzberichte...\n");

    let languages = [
        "Deutsch",
        "Englisch",
        "Französisch",
        "Spanisch",
        "Portugiesisch",
    ];

    for lang_key in &languages {
        let config = LANG_CONFIG
            .get(*lang_key)
            .expect("Language config not found");
        let suffix = LANG_SUFFIXES
            .get(*lang_key)
            .expect("Language suffix not found");

        println!("Erstelle {} ({})...", lang_key, config.lang_val);

        let mut workbook = Workbook::new();
        let styles = ReportStyles::new();

        // Sprachversionen-Sheet hinzufügen (für VLOOKUP-Formeln)
        build_sprachversionen(&mut workbook)?;

        // Worksheet erstellen mit lokalem Namen
        let ws = workbook.add_worksheet();
        ws.set_name(config.fb_sheet)?;

        // Alle Spalten standardmäßig unlocked setzen
        let unlocked = Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_unlocked();
        for col in 0..1000u16 {
            ws.set_column_format(col, &unlocked).ok();
        }

        setup_sheet(ws)?;

        // Test-Werte setzen
        let mut values = ReportValues::new();

        // WICHTIG: Sprache setzen (wird in E2 geschrieben für VLOOKUP)
        values.set(ApiKey::Language, config.lang_val);
        values.set(ApiKey::Currency, "EUR");
        values.set(
            ApiKey::ProjectNumber,
            format!(
                "PROJ-2024-{}",
                suffix.trim_start_matches('_').to_uppercase()
            ),
        );
        values.set(
            ApiKey::ProjectTitle,
            format!("Test Project ({})", config.lang_val),
        );
        values.set(ApiKey::ProjectStart, "01.01.2024");
        values.set(ApiKey::ProjectEnd, "31.12.2024");
        values.set(ApiKey::ReportStart, "01.01.2024");
        values.set(ApiKey::ReportEnd, "30.06.2024");

        // Einnahmen-Tabelle
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), 10000.0 * (i + 1) as f64);
            values.set(ApiKey::IncomeReportPeriod(i), 5000.0 * (i + 1) as f64);
            values.set(ApiKey::IncomeTotal(i), 5000.0 * (i + 1) as f64);
        }

        // Body-Konfiguration
        let body_config = BodyConfig::new()
            .with_positions(1, 5)
            .with_positions(2, 3)
            .with_positions(3, 4)
            .with_positions(4, 3)
            .with_positions(5, 2)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);

        // Kostenpositionen
        values.set_position_row(1, 1, "Position 1.1", 15000.0, 7500.0, 7500.0, "");
        values.set_position_row(1, 2, "Position 1.2", 25000.0, 12000.0, 12000.0, "");
        values.set_position_row(1, 3, "Position 1.3", 5000.0, 2500.0, 2500.0, "");
        values.set_position_row(1, 4, "Position 1.4", 8000.0, 4000.0, 4000.0, "");
        values.set_position_row(1, 5, "Position 1.5", 2000.0, 1000.0, 1000.0, "");

        values.set_position_row(2, 1, "Position 2.1", 1500.0, 800.0, 800.0, "");
        values.set_position_row(2, 2, "Position 2.2", 3000.0, 2000.0, 2000.0, "");
        values.set_position_row(2, 3, "Position 2.3", 2000.0, 1500.0, 1500.0, "");

        values.set_position_row(3, 1, "Position 3.1", 2000.0, 1200.0, 1200.0, "");
        values.set_position_row(3, 2, "Position 3.2", 5000.0, 2500.0, 2500.0, "");
        values.set_position_row(3, 3, "Position 3.3", 1000.0, 500.0, 500.0, "");
        values.set_position_row(3, 4, "Position 3.4", 800.0, 400.0, 400.0, "");

        values.set_position_row(4, 1, "Position 4.1", 10000.0, 5000.0, 5000.0, "");
        values.set_position_row(4, 2, "Position 4.2", 8000.0, 4000.0, 4000.0, "");
        values.set_position_row(4, 3, "Position 4.3", 2000.0, 1000.0, 1000.0, "");

        values.set_position_row(5, 1, "Position 5.1", 3000.0, 1500.0, 1500.0, "");
        values.set_position_row(5, 2, "Position 5.2", 1500.0, 750.0, 750.0, "");

        // Header-Eingabe Kategorien
        values.set_header_input(6, 4000.0, 2000.0, 2000.0, "");
        values.set_header_input(7, 6000.0, 3000.0, 3000.0, "");
        values.set_header_input(8, 2500.0, 1250.0, 1250.0, "");

        // Footer-Werte (Saldenabstimmung)
        values.set_footer_bank(10000.0);
        values.set_footer_kasse(2000.0);
        values.set_footer_sonstiges(500.0);

        // Report schreiben
        write_report_v2_with_body(ws, &styles, suffix, &values, &body_config)?;

        // Blattschutz aktivieren
        ws.protect();

        // Datei speichern
        let filename = format!("{}.xlsx", config.base);
        workbook.save(&filename)?;
        println!("  -> {} gespeichert", filename);
    }

    println!("\nAlle 5 Dateien erstellt:");
    for lang_key in &languages {
        let config = LANG_CONFIG.get(*lang_key).unwrap();
        println!("  - {}.xlsx ({})", config.base, config.lang_val);
    }

    Ok(())
}
