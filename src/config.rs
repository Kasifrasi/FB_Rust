//! Einheitliche Report-Konfiguration für Tauri-Integration
//!
//! Dieser Modul stellt einen einfachen, serialisierbaren Wrapper bereit,
//! der alle Parameter für einen Finanzbericht in einem einzigen Struct bündelt.
//!
//! ## Tauri-Beispiel
//!
//! ```ignore
//! #[tauri::command]
//! fn generate_report(config: ReportConfig, output_path: String) -> Result<(), String> {
//!     config.write_to(output_path).map_err(|e| e.to_string())
//! }
//! ```
//!
//! ## TypeScript-Beispiel
//!
//! ```typescript
//! await invoke("generate_report", {
//!     config: {
//!         language: "deutsch",
//!         currency: "EUR",
//!         project_number: "2025-001",
//!         project_title: "Klimaschutzprojekt",
//!         locked: true,
//!         body_positions: { "1": 10, "2": 5 },
//!         positions: [
//!             { category: 1, position: 1, description: "Baukosten", approved: 50000.0 }
//!         ],
//!         footer_bank: 12000.0,
//!         footer_kasse: 500.0,
//!     },
//!     outputPath: "/home/user/report.xlsx"
//! });
//! ```

use std::collections::HashMap;
use std::path::Path;

use crate::report::api::{ApiKey, CellValue, ReportValues};
use crate::report::body::BodyConfig;
use crate::report::format::{ReportOptions, ReportStyles};
use crate::report::writer::create_protected_report;

// ============================================================================
// Hilfstypen
// ============================================================================

/// Einzelne Zeile des oberen Tabellen-Bereichs (Zeilen 15-19, Index 0-4)
///
/// Enthält bewilligtes Budget und Einnahmen-Felder für eine der 5 Tabellenzeilen.
#[derive(Debug, Clone)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct TableEntry {
    /// 0-basierter Index (0-4)
    pub index: u8,
    /// Spalte D: Bewilligtes Budget
    pub approved_budget: Option<f64>,
    /// Spalte E: Einnahmen Berichtszeitraum
    pub income_report: Option<f64>,
    /// Spalte F: Einnahmen gesamt
    pub income_total: Option<f64>,
    /// Spalte H: Begründung
    pub reason: Option<String>,
}

/// Einzelne Zeile eines Kassenbuch-Panels (links oder rechts, Index 0-17)
#[derive(Debug, Clone)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct PanelEntry {
    /// 0-basierter Index (0-17)
    pub index: u8,
    /// Datum als String (z.B. "2025-01-15" oder "15.01.2025")
    pub date: Option<String>,
    /// Betrag in Euro
    pub amount_euro: Option<f64>,
    /// Betrag in Lokalwährung
    pub amount_local: Option<f64>,
}

/// Einzelne Kostenpositions-Zeile im Body-Bereich
#[derive(Debug, Clone)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct PositionEntry {
    /// Kategorie (1-8)
    pub category: u8,
    /// Positions-Nummer: 0 = Header-Eingabe-Modus, 1..N = normale Position
    pub position: u16,
    /// Spalte C: Beschreibung (nur bei position >= 1)
    pub description: Option<String>,
    /// Spalte D: Bewilligtes Budget
    pub approved: Option<f64>,
    /// Spalte E: Einnahmen Berichtszeitraum
    pub income_report: Option<f64>,
    /// Spalte F: Einnahmen gesamt
    pub income_total: Option<f64>,
    /// Spalte H: Begründung/Bemerkung
    pub remark: Option<String>,
}

// ============================================================================
// ReportConfig — Haupt-Wrapper
// ============================================================================

/// Einheitliche Konfiguration für einen Finanzbericht
///
/// Bündelt alle Parameter in einem einzigen, serialisierbaren Struct.
/// Primär für die Tauri-Integration gedacht, aber auch direkt in Rust verwendbar.
///
/// `ReportStyles` ist absichtlich nicht Teil dieses Structs — Styles sind
/// stabil und werden immer mit den Standard-Einstellungen erzeugt.
///
/// ## Beispiel
///
/// ```ignore
/// use kmw_fb_rust::ReportConfig;
///
/// let config = ReportConfig {
///     language: "deutsch".to_string(),
///     currency: "EUR".to_string(),
///     locked: true,
///     body_positions: [(1, 10), (2, 5)].into(),
///     ..ReportConfig::default()
/// };
/// config.write_to("output/report.xlsx").unwrap();
/// ```
#[derive(Debug, Clone)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct ReportConfig {
    // -------------------------------------------------------------------------
    // Header-Felder
    // -------------------------------------------------------------------------
    /// Sprache: "deutsch", "english", "francais", "espanol", "portugues"
    pub language: String,
    /// Währungscode (ISO 4217), z.B. "EUR", "USD"
    pub currency: String,
    /// Projektnummer (Zelle D5)
    pub project_number: Option<String>,
    /// Projekttitel (Zelle D6)
    pub project_title: Option<String>,
    /// Projektstart-Datum (Zelle E8)
    pub project_start: Option<String>,
    /// Projektende-Datum (Zelle G8)
    pub project_end: Option<String>,
    /// Berichtszeitraum Beginn (Zelle E9)
    pub report_start: Option<String>,
    /// Berichtszeitraum Ende (Zelle G9)
    pub report_end: Option<String>,

    // -------------------------------------------------------------------------
    // Tabellen-Bereich (Zeilen 15-19)
    // -------------------------------------------------------------------------
    /// Einträge für den oberen Tabellen-Bereich (max. 5 Zeilen, Index 0-4)
    #[cfg_attr(feature = "serde", serde(default))]
    pub table: Vec<TableEntry>,

    // -------------------------------------------------------------------------
    // Kassenbuch-Panels
    // -------------------------------------------------------------------------
    /// Linkes Kassenbuch-Panel (max. 18 Zeilen, Index 0-17)
    #[cfg_attr(feature = "serde", serde(default))]
    pub left_panel: Vec<PanelEntry>,
    /// Rechtes Kassenbuch-Panel (max. 18 Zeilen, Index 0-17)
    #[cfg_attr(feature = "serde", serde(default))]
    pub right_panel: Vec<PanelEntry>,

    // -------------------------------------------------------------------------
    // Body-Kostenpositionen
    // -------------------------------------------------------------------------
    /// Kostenpositionen im Body-Bereich
    #[cfg_attr(feature = "serde", serde(default))]
    pub positions: Vec<PositionEntry>,
    /// Anzahl Positions-Zeilen pro Kategorie (1-8)
    ///
    /// Schlüssel = Kategorie-Nummer, Wert = Anzahl Zeilen (0 = Header-Eingabe-Modus)
    pub body_positions: HashMap<u8, u16>,

    // -------------------------------------------------------------------------
    // Footer-Salden
    // -------------------------------------------------------------------------
    /// Bank-Saldo für Saldenabstimmung
    pub footer_bank: Option<f64>,
    /// Kassen-Saldo für Saldenabstimmung
    pub footer_kasse: Option<f64>,
    /// Sonstige Salden (Schecks, Vorschüsse, etc.)
    pub footer_sonstiges: Option<f64>,

    // -------------------------------------------------------------------------
    // Optionen
    // -------------------------------------------------------------------------
    /// Sheet sperren (Standard-Schutz für alle Zellen außer Eingabefeldern)
    pub locked: bool,
    /// Passwort für Workbook-Schutz (`None` = kein Workbook-Schutz)
    ///
    /// Verhindert das Hinzufügen, Löschen und Umbenennen von Sheets.
    pub workbook_password: Option<String>,
}

impl Default for ReportConfig {
    fn default() -> Self {
        Self {
            language: "deutsch".to_string(),
            currency: "EUR".to_string(),
            project_number: None,
            project_title: None,
            project_start: None,
            project_end: None,
            report_start: None,
            report_end: None,
            table: Vec::new(),
            left_panel: Vec::new(),
            right_panel: Vec::new(),
            positions: Vec::new(),
            body_positions: BodyConfig::default_positions(),
            footer_bank: None,
            footer_kasse: None,
            footer_sonstiges: None,
            locked: false,
            workbook_password: None,
        }
    }
}

impl ReportConfig {
    /// Schreibt den Finanzbericht in die angegebene Datei.
    ///
    /// Styles werden intern mit Standard-Einstellungen erzeugt.
    pub fn write_to(&self, output_path: impl AsRef<Path>) -> Result<(), Box<dyn std::error::Error>> {
        let values = self.build_values();
        let body_config = self.build_body_config();
        let options = self.build_options();
        let styles = ReportStyles::new();
        create_protected_report(output_path, &styles, &values, &body_config, &options)?;
        Ok(())
    }

    fn build_values(&self) -> ReportValues {
        let mut v = ReportValues::new()
            .with_language(&self.language)
            .with_currency(&self.currency);

        if let Some(ref s) = self.project_number {
            v = v.with_project_number(s);
        }
        if let Some(ref s) = self.project_title {
            v = v.with_project_title(s);
        }
        if let Some(ref s) = self.project_start {
            v = v.with_project_start(s);
        }
        if let Some(ref s) = self.project_end {
            v = v.with_project_end(s);
        }
        if let Some(ref s) = self.report_start {
            v = v.with_report_start(s);
        }
        if let Some(ref s) = self.report_end {
            v = v.with_report_end(s);
        }

        // Tabelle (Zeilen 15-19)
        for e in &self.table {
            if let Some(n) = e.approved_budget {
                v.set(ApiKey::ApprovedBudget(e.index), n);
            }
            if let Some(n) = e.income_report {
                v.set(ApiKey::IncomeReportPeriod(e.index), n);
            }
            if let Some(n) = e.income_total {
                v.set(ApiKey::IncomeTotal(e.index), n);
            }
            if let Some(ref s) = e.reason {
                v.set(ApiKey::IncomeReason(e.index), s.as_str());
            }
        }

        // Linkes Panel
        for e in &self.left_panel {
            if let Some(ref s) = e.date {
                v.set(ApiKey::LeftDate(e.index), CellValue::Date(s.clone()));
            }
            if let Some(n) = e.amount_euro {
                v.set(ApiKey::LeftAmountEuro(e.index), n);
            }
            if let Some(n) = e.amount_local {
                v.set(ApiKey::LeftAmountLocal(e.index), n);
            }
        }

        // Rechtes Panel
        for e in &self.right_panel {
            if let Some(ref s) = e.date {
                v.set(ApiKey::RightDate(e.index), CellValue::Date(s.clone()));
            }
            if let Some(n) = e.amount_euro {
                v.set(ApiKey::RightAmountEuro(e.index), n);
            }
            if let Some(n) = e.amount_local {
                v.set(ApiKey::RightAmountLocal(e.index), n);
            }
        }

        // Body-Kostenpositionen
        for e in &self.positions {
            if e.position == 0 {
                // Header-Eingabe-Modus
                v.set_header_input(
                    e.category,
                    opt_num(e.approved),
                    opt_num(e.income_report),
                    opt_num(e.income_total),
                    opt_str(e.remark.clone()),
                );
            } else {
                v.set_position_row(
                    e.category,
                    e.position,
                    opt_str(e.description.clone()),
                    opt_num(e.approved),
                    opt_num(e.income_report),
                    opt_num(e.income_total),
                    opt_str(e.remark.clone()),
                );
            }
        }

        // Footer-Salden
        if let Some(n) = self.footer_bank {
            v.set_footer_bank(n);
        }
        if let Some(n) = self.footer_kasse {
            v.set_footer_kasse(n);
        }
        if let Some(n) = self.footer_sonstiges {
            v.set_footer_sonstiges(n);
        }

        v
    }

    fn build_body_config(&self) -> BodyConfig {
        let mut config = BodyConfig::new();
        for (&cat, &count) in &self.body_positions {
            config = config.with_positions(cat, count);
        }
        config
    }

    fn build_options(&self) -> ReportOptions {
        let mut opts = if self.locked {
            ReportOptions::with_default_protection()
        } else {
            ReportOptions::new()
        };
        if let Some(ref pw) = self.workbook_password {
            opts = opts.with_workbook_protection(pw);
        }
        opts
    }
}

// ============================================================================
// Hilfsfunktionen
// ============================================================================

fn opt_num(v: Option<f64>) -> CellValue {
    match v {
        Some(n) => CellValue::Number(n),
        None => CellValue::Empty,
    }
}

fn opt_str(v: Option<String>) -> CellValue {
    match v {
        Some(s) if !s.is_empty() => CellValue::Text(s),
        _ => CellValue::Empty,
    }
}
