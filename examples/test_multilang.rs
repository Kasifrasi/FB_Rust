//! Test: Mehrsprachige Finanzberichte
//!
//! Generiert 5 Excel-Dateien in verschiedenen Sprachen:
//! - Deutsch (D_Finanzbericht.xlsx)
//! - Englisch (E_Financial_report.xlsx)
//! - Französisch (F_Rapport_financier.xlsx)
//! - Spanisch (S_Informe_financiero.xlsx)
//! - Portugiesisch (P_Relatorio_financeiro.xlsx)

use fb_rust::{PositionEntry, ReportConfig, TableEntry};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle mehrsprachige Finanzberichte...\n");

    let reports: &[(&str, &str, &str)] = &[
        ("deutsch",   "PROJ-2024-DE", "examples/output/D_Finanzbericht.xlsx"),
        ("english",   "PROJ-2024-EN", "examples/output/E_Financial_report.xlsx"),
        ("francais",  "PROJ-2024-FR", "examples/output/F_Rapport_financier.xlsx"),
        ("espanol",   "PROJ-2024-ES", "examples/output/S_Informe_financiero.xlsx"),
        ("portugues", "PROJ-2024-PT", "examples/output/P_Relatorio_financeiro.xlsx"),
    ];

    for (language, proj_nr, filename) in reports {
        create_report(language, proj_nr, filename)?;
        println!("  -> {} gespeichert", filename);
    }

    println!("\nAlle 5 Dateien erstellt.");
    Ok(())
}

fn create_report(
    language: &str,
    project_number: &str,
    output_path: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Erstelle {}...", language);

    let config = ReportConfig {
        language: language.to_string(),
        currency: "EUR".to_string(),
        project_number: Some(project_number.to_string()),
        project_title: Some(format!("Test Project ({})", language)),
        project_start: Some("01.01.2024".to_string()),
        project_end: Some("31.12.2024".to_string()),
        report_start: Some("01.01.2024".to_string()),
        report_end: Some("30.06.2024".to_string()),

        table: (0..5u8)
            .map(|i| TableEntry {
                index: i,
                approved_budget: Some(10000.0 * (i + 1) as f64),
                income_report: Some(5000.0 * (i + 1) as f64),
                income_total: Some(5000.0 * (i + 1) as f64),
                reason: None,
            })
            .collect(),

        positions: vec![
            PositionEntry { category: 1, position: 1, description: Some("Position 1.1".to_string()), approved: Some(15000.0), income_report: Some(7500.0), income_total: Some(7500.0), remark: None },
            PositionEntry { category: 1, position: 2, description: Some("Position 1.2".to_string()), approved: Some(25000.0), income_report: Some(12000.0), income_total: Some(12000.0), remark: None },
            PositionEntry { category: 1, position: 3, description: Some("Position 1.3".to_string()), approved: Some(5000.0), income_report: Some(2500.0), income_total: Some(2500.0), remark: None },
            PositionEntry { category: 1, position: 4, description: Some("Position 1.4".to_string()), approved: Some(8000.0), income_report: Some(4000.0), income_total: Some(4000.0), remark: None },
            PositionEntry { category: 1, position: 5, description: Some("Position 1.5".to_string()), approved: Some(2000.0), income_report: Some(1000.0), income_total: Some(1000.0), remark: None },
            PositionEntry { category: 2, position: 1, description: Some("Position 2.1".to_string()), approved: Some(1500.0), income_report: Some(800.0), income_total: Some(800.0), remark: None },
            PositionEntry { category: 2, position: 2, description: Some("Position 2.2".to_string()), approved: Some(3000.0), income_report: Some(2000.0), income_total: Some(2000.0), remark: None },
            PositionEntry { category: 2, position: 3, description: Some("Position 2.3".to_string()), approved: Some(2000.0), income_report: Some(1500.0), income_total: Some(1500.0), remark: None },
            PositionEntry { category: 3, position: 1, description: Some("Position 3.1".to_string()), approved: Some(2000.0), income_report: Some(1200.0), income_total: Some(1200.0), remark: None },
            PositionEntry { category: 3, position: 2, description: Some("Position 3.2".to_string()), approved: Some(5000.0), income_report: Some(2500.0), income_total: Some(2500.0), remark: None },
            PositionEntry { category: 3, position: 3, description: Some("Position 3.3".to_string()), approved: Some(1000.0), income_report: Some(500.0), income_total: Some(500.0), remark: None },
            PositionEntry { category: 3, position: 4, description: Some("Position 3.4".to_string()), approved: Some(800.0), income_report: Some(400.0), income_total: Some(400.0), remark: None },
            PositionEntry { category: 4, position: 1, description: Some("Position 4.1".to_string()), approved: Some(10000.0), income_report: Some(5000.0), income_total: Some(5000.0), remark: None },
            PositionEntry { category: 4, position: 2, description: Some("Position 4.2".to_string()), approved: Some(8000.0), income_report: Some(4000.0), income_total: Some(4000.0), remark: None },
            PositionEntry { category: 4, position: 3, description: Some("Position 4.3".to_string()), approved: Some(2000.0), income_report: Some(1000.0), income_total: Some(1000.0), remark: None },
            PositionEntry { category: 5, position: 1, description: Some("Position 5.1".to_string()), approved: Some(3000.0), income_report: Some(1500.0), income_total: Some(1500.0), remark: None },
            PositionEntry { category: 5, position: 2, description: Some("Position 5.2".to_string()), approved: Some(1500.0), income_report: Some(750.0), income_total: Some(750.0), remark: None },
            // Header-Eingabe Kategorien (position = 0)
            PositionEntry { category: 6, position: 0, description: None, approved: Some(4000.0), income_report: Some(2000.0), income_total: Some(2000.0), remark: None },
            PositionEntry { category: 7, position: 0, description: None, approved: Some(6000.0), income_report: Some(3000.0), income_total: Some(3000.0), remark: None },
            PositionEntry { category: 8, position: 0, description: None, approved: Some(2500.0), income_report: Some(1250.0), income_total: Some(1250.0), remark: None },
        ],

        body_positions: [(1u8, 5u16), (2, 3), (3, 4), (4, 3), (5, 2), (6, 0), (7, 0), (8, 0)]
            .into_iter()
            .collect(),

        footer_bank: Some(10000.0),
        footer_kasse: Some(2000.0),
        footer_sonstiges: Some(500.0),

        locked: true,
        hide_columns_qv: true,

        ..ReportConfig::default()
    };

    config.write_to(output_path)?;
    Ok(())
}
