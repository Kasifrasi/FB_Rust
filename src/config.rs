//! Unified report configuration for Tauri integration.
//!
//! This module provides a single, serializable struct ([`ReportConfig`]) that bundles
//! all parameters for a financial report. It is the recommended entry point for both
//! Rust and Tauri/JSON usage.
//!
//! ## Tauri Example
//!
//! ```ignore
//! #[tauri::command]
//! fn generate_report(config: ReportConfig, output_path: String) -> Result<(), String> {
//!     config.write_to(output_path).map_err(|e| e.to_string())
//! }
//! ```
//!
//! ## TypeScript Example
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
use crate::report::options::{ReportOptions, RowGrouping};
use crate::report::writer::{create_protected_report, create_protected_report_precomputed};
use crate::workbook_protection::WorkbookProtection;

// ============================================================================
// Hilfstypen
// ============================================================================

/// A single row in the upper table area (Excel rows 15-19, index 0-4).
///
/// Contains approved budget and income fields for one of the 5 table rows.
///
/// ## JSON
///
/// ```json
/// { "index": 0, "approved_budget": 50000.0, "income_report": null, "income_total": null, "reason": "Spende" }
/// ```
#[derive(Debug, Clone, PartialEq)]
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

/// A single row in a cash book panel (left or right, index 0-17).
///
/// ## JSON
///
/// ```json
/// { "index": 0, "date": "2025-01-15", "amount_euro": 1200.50, "amount_local": null }
/// ```
#[derive(Debug, Clone, PartialEq)]
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

/// A single cost position row in the body area.
///
/// ## JSON
///
/// Normal position (`position >= 1`):
/// ```json
/// { "category": 1, "position": 1, "description": "Personalkosten", "approved": 50000.0, "income_report": null, "income_total": null, "remark": null }
/// ```
///
/// Header-input mode (`position == 0`, category has 0 rows in `body_positions`):
/// ```json
/// { "category": 6, "position": 0, "description": null, "approved": 3000.0, "income_report": 1500.0, "income_total": 1500.0, "remark": "Sonstiges" }
/// ```
#[derive(Debug, Clone, PartialEq)]
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

/// Unified configuration for a financial report.
///
/// Bundles all parameters in a single, serializable struct. Primary entry point
/// for both Tauri/JSON integration and direct Rust usage. Styles are generated
/// internally with fixed defaults.
///
/// ## Rust Example
///
/// ```ignore
/// use fb_rust::ReportConfig;
///
/// let config = ReportConfig {
///     language: "deutsch".to_string(),
///     currency: "EUR".to_string(),
///     locked: true,
///     body_positions: [(1, 10), (2, 5)].into(),
///     ..ReportConfig::default()
/// };
/// config.write_to("report.xlsx").unwrap();
/// ```
///
/// ## JSON (minimal)
///
/// ```json
/// {
///   "language": "deutsch",
///   "currency": "EUR",
///   "body_positions": { "1": 20, "2": 20, "3": 30, "4": 30, "5": 20, "6": 0, "7": 0, "8": 0 },
///   "locked": false,
///   "hide_columns_qv": false,
///   "hide_language_sheet": false
/// }
/// ```
///
/// ## JSON (full)
///
/// ```json
/// {
///   "language": "english",
///   "currency": "USD",
///   "project_number": "PROJ-2025-001",
///   "project_title": "Education Project",
///   "project_start": "01.01.2025",
///   "project_end": "31.12.2025",
///   "report_start": "01.01.2025",
///   "report_end": "30.06.2025",
///   "table": [
///     { "index": 0, "approved_budget": 50000.0, "income_report": 25000.0, "income_total": 25000.0, "reason": "Donation" }
///   ],
///   "left_panel": [
///     { "index": 0, "date": "15.01.2025", "amount_euro": 1000.0, "amount_local": 1100.0 }
///   ],
///   "right_panel": [],
///   "positions": [
///     { "category": 1, "position": 1, "description": "Personnel", "approved": 18000.0, "income_report": 9000.0, "income_total": 9000.0, "remark": null }
///   ],
///   "body_positions": { "1": 10, "2": 5, "6": 0 },
///   "footer_bank": 8500.0,
///   "footer_kasse": 250.50,
///   "footer_sonstiges": null,
///   "locked": true,
///   "workbook_password": "secret",
///   "hide_columns_qv": true,
///   "hide_language_sheet": true,
///   "row_grouping": {
///     "groups": [{ "start_row": 10, "end_row": 20, "collapsed": false }],
///     "symbols_above": false
///   }
/// }
/// ```
///
/// ## Field Constraints
///
/// | Field | Constraint |
/// |-------|-----------|
/// | `language` | `"deutsch"`, `"english"`, `"francais"`, `"espanol"`, `"portugues"` |
/// | `currency` | ISO 4217 code (e.g. `"EUR"`, `"USD"`), validated at report write time |
/// | `table` | Max 5 entries, `index` 0–4 |
/// | `left_panel` / `right_panel` | Max 18 entries each, `index` 0–17 |
/// | `positions` | `category` 1–8, `position` 0 = header-input, 1..N = row |
/// | `body_positions` | Keys `"1"` – `"8"`, values = row count (0 = header-input) |
///
/// ## Errors
///
/// See [`ReportError`](crate::ReportError) for error variants returned by
/// [`write_to`](Self::write_to) and [`write_to_precomputed`](Self::write_to_precomputed).
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
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
    /// Hilfsspalten Q-V ausblenden
    pub hide_columns_qv: bool,
    /// Language-Sheet ausblenden
    pub hide_language_sheet: bool,
    /// Zeilen-Gruppierung für ausklappbare Abschnitte
    #[cfg_attr(feature = "serde", serde(default))]
    pub row_grouping: Option<RowGrouping>,
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
            hide_columns_qv: false,
            hide_language_sheet: false,
            row_grouping: None,
        }
    }
}

impl ReportConfig {
    /// Writes the financial report to the given file path.
    ///
    /// # Errors
    ///
    /// Returns [`ReportError`](crate::ReportError):
    /// - [`Xlsx`](crate::ReportError::Xlsx) — Excel write failure
    /// - [`Protection`](crate::ReportError::Protection) — workbook password injection failed (only with `workbook_password`)
    /// - [`Io`](crate::ReportError::Io) — file system error
    /// - [`InvalidPath`](crate::ReportError::InvalidPath) — path contains non-UTF-8
    pub fn write_to(&self, output_path: impl AsRef<Path>) -> Result<(), crate::error::ReportError> {
        let values = self.build_values();
        let body_config = self.build_body_config();
        let options = self.build_options();
        let wb_prot = self.workbook_password.as_ref().map(WorkbookProtection::new);
        create_protected_report(
            output_path, &values, &body_config, &options,
            wb_prot.as_ref(), self.hide_language_sheet,
        )?;
        Ok(())
    }

    /// Batch-optimized: precompute the hash once, reuse for N files.
    ///
    /// Saves ~25ms SHA-512 overhead per file when using the same password.
    /// The `workbook_password` field is ignored — the precomputed hash is used directly.
    ///
    /// ```ignore
    /// use fb_rust::{precompute_hash, ReportConfig};
    ///
    /// let hash = precompute_hash("password");
    /// for config in &configs {
    ///     config.write_to_precomputed("output.xlsx", &hash).unwrap();
    /// }
    /// ```
    ///
    /// # Errors
    ///
    /// Same as [`write_to`](Self::write_to). See [`ReportError`](crate::ReportError).
    pub fn write_to_precomputed(
        &self,
        output_path: impl AsRef<Path>,
        hash: &crate::workbook_protection::PrecomputedHash,
    ) -> Result<(), crate::error::ReportError> {
        let values = self.build_values();
        let body_config = self.build_body_config();
        let options = self.build_options();
        create_protected_report_precomputed(
            output_path, &values, &body_config, &options,
            self.hide_language_sheet, hash,
        )?;
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
        if self.hide_columns_qv {
            opts = opts.with_hidden_columns_qv();
        }
        if let Some(ref rg) = self.row_grouping {
            opts = opts.with_row_grouping(rg.clone());
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
