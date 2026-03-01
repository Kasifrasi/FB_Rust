//! Unified report configuration for Tauri integration.
//!
//! This module provides [`ReportConfig`] as the main entry point, composed of
//! four domain sub-structs: [`ReportHeader`], [`ReportBody`], [`ReportFooter`],
//! and [`ReportOptions`].
//!
//! All structs support the Builder pattern via `derive_builder`.
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
//!         header: {
//!             language: "deutsch",
//!             currency: "EUR",
//!             project_number: "2025-001",
//!             project_title: "Klimaschutzprojekt",
//!         },
//!         body: {
//!             body_positions: { "1": 10, "2": 5 },
//!             positions: [
//!                 { category: 1, position: 1, description: "Baukosten", approved: 50000.0 }
//!             ],
//!         },
//!         footer: { bank: 12000.0, kasse: 500.0 },
//!         options: { sheet_password: "geheim" },
//!     },
//!     outputPath: "/home/user/report.xlsx"
//! });
//! ```

use std::collections::HashMap;
use std::path::Path;

use derive_builder::Builder;

use crate::report::api::{ApiKey, CellValue, Currency, Language, ReportValues};
use crate::report::body::BodyConfig;
use crate::report::options::{RowGrouping, SheetOptions, SheetProtection};
use crate::report::writer::{create_protected_report, create_protected_report_precomputed};
use crate::workbook_protection::WorkbookProtection;

// ============================================================================
// Entry types (TableEntry, PanelEntry, PositionEntry)
// ============================================================================

/// A single row in the upper table area (Excel rows 15-19, index 0-4).
///
/// ## JSON
///
/// ```json
/// { "index": 0, "approved_budget": 50000.0, "income_report": null, "income_total": null, "reason": "Spende" }
/// ```
#[derive(Debug, Clone, PartialEq, Default, Builder)]
#[builder(setter(into, strip_option), default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct TableEntry {
    /// 0-based index (0-4)
    pub index: u8,
    /// Column D: Approved budget
    pub approved_budget: Option<f64>,
    /// Column E: Income report period
    pub income_report: Option<f64>,
    /// Column F: Income total
    pub income_total: Option<f64>,
    /// Column H: Reason
    pub reason: Option<String>,
}

/// A single row in a cash book panel (left or right, index 0-17).
///
/// ## JSON
///
/// ```json
/// { "index": 0, "date": "2025-01-15", "amount_euro": 1200.50, "amount_local": null }
/// ```
#[derive(Debug, Clone, PartialEq, Default, Builder)]
#[builder(setter(into, strip_option), default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct PanelEntry {
    /// 0-based index (0-17)
    pub index: u8,
    /// Date string (e.g. "2025-01-15" or "15.01.2025")
    pub date: Option<String>,
    /// Amount in Euro
    pub amount_euro: Option<f64>,
    /// Amount in local currency
    pub amount_local: Option<f64>,
}

/// A single cost position row in the body area.
///
/// ## JSON
///
/// Normal position (`position >= 1`):
/// ```json
/// { "category": 1, "position": 1, "description": "Personalkosten", "approved": 50000.0 }
/// ```
///
/// Header-input mode (`position == 0`, category has 0 rows in `body_positions`):
/// ```json
/// { "category": 6, "position": 0, "approved": 3000.0, "remark": "Sonstiges" }
/// ```
#[derive(Debug, Clone, PartialEq, Default, Builder)]
#[builder(setter(into, strip_option), default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub struct PositionEntry {
    /// Category (1-8)
    pub category: u8,
    /// Position number: 0 = header-input mode, 1..N = normal position
    pub position: u16,
    /// Column C: Description (only for position >= 1)
    pub description: Option<String>,
    /// Column D: Approved budget
    pub approved: Option<f64>,
    /// Column E: Income report period
    pub income_report: Option<f64>,
    /// Column F: Income total
    pub income_total: Option<f64>,
    /// Column H: Remark
    pub remark: Option<String>,
}

// ============================================================================
// ReportHeader
// ============================================================================

/// Header metadata for a financial report.
///
/// Contains language, currency, project info, and date ranges.
///
/// ## JSON
///
/// ```json
/// {
///   "language": "deutsch",
///   "currency": "EUR",
///   "project_number": "PROJ-2025-001",
///   "project_title": "Bildungsprojekt",
///   "project_start": "01.01.2025",
///   "project_end": "31.12.2025",
///   "report_start": "01.01.2025",
///   "report_end": "30.06.2025"
/// }
/// ```
#[derive(Debug, Clone, PartialEq, Builder)]
#[builder(default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportHeader {
    /// Language (validated enum)
    pub language: Language,
    /// Currency (validated ISO 4217 code)
    pub currency: Currency,
    /// Project number (cell D5)
    #[builder(setter(into, strip_option))]
    pub project_number: Option<String>,
    /// Project title (cell D6)
    #[builder(setter(into, strip_option))]
    pub project_title: Option<String>,
    /// Project start date (cell E8)
    #[builder(setter(into, strip_option))]
    pub project_start: Option<String>,
    /// Project end date (cell G8)
    #[builder(setter(into, strip_option))]
    pub project_end: Option<String>,
    /// Report period start (cell E9)
    #[builder(setter(into, strip_option))]
    pub report_start: Option<String>,
    /// Report period end (cell G9)
    #[builder(setter(into, strip_option))]
    pub report_end: Option<String>,
}

impl Default for ReportHeader {
    fn default() -> Self {
        Self {
            language: Language::Deutsch,
            currency: Currency::EUR,
            project_number: None,
            project_title: None,
            project_start: None,
            project_end: None,
            report_start: None,
            report_end: None,
        }
    }
}

// ============================================================================
// ReportBody
// ============================================================================

/// Body data: income table, cash book panels, and cost positions.
///
/// ## JSON
///
/// ```json
/// {
///   "table": [{ "index": 0, "approved_budget": 50000.0 }],
///   "left_panel": [{ "index": 0, "date": "15.01.2025", "amount_euro": 1000.0 }],
///   "right_panel": [],
///   "positions": [{ "category": 1, "position": 1, "description": "Personnel", "approved": 18000.0 }],
///   "body_positions": { "1": 10, "2": 5, "6": 0 }
/// }
/// ```
#[derive(Debug, Clone, PartialEq, Builder)]
#[builder(default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportBody {
    /// Income table rows (max 5, index 0-4)
    #[cfg_attr(feature = "serde", serde(default))]
    #[builder(setter(each(name = "add_table_entry")))]
    pub table: Vec<TableEntry>,
    /// Left cash book panel (max 18 rows, index 0-17)
    #[cfg_attr(feature = "serde", serde(default))]
    #[builder(setter(each(name = "add_left_panel_entry")))]
    pub left_panel: Vec<PanelEntry>,
    /// Right cash book panel (max 18 rows, index 0-17)
    #[cfg_attr(feature = "serde", serde(default))]
    #[builder(setter(each(name = "add_right_panel_entry")))]
    pub right_panel: Vec<PanelEntry>,
    /// Cost positions in the body area
    #[cfg_attr(feature = "serde", serde(default))]
    #[builder(setter(each(name = "add_position")))]
    pub positions: Vec<PositionEntry>,
    /// Number of position rows per category (1-8).
    /// Key = category number, value = row count (0 = header-input mode)
    pub body_positions: HashMap<u8, u16>,
}

impl Default for ReportBody {
    fn default() -> Self {
        Self {
            table: Vec::new(),
            left_panel: Vec::new(),
            right_panel: Vec::new(),
            positions: Vec::new(),
            body_positions: BodyConfig::default_positions(),
        }
    }
}

// ============================================================================
// ReportFooter
// ============================================================================

/// Footer balances for the reconciliation section.
///
/// ## JSON
///
/// ```json
/// { "bank": 8500.0, "kasse": 250.50, "sonstiges": null }
/// ```
#[derive(Debug, Clone, PartialEq, Default, Builder)]
#[builder(setter(strip_option), default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportFooter {
    /// Bank balance
    pub bank: Option<f64>,
    /// Cash balance
    pub kasse: Option<f64>,
    /// Other balances (checks, advances, etc.)
    pub sonstiges: Option<f64>,
}

// ============================================================================
// ReportOptions (user-facing)
// ============================================================================

/// Output options: protection, visibility, and row grouping.
///
/// ## JSON
///
/// ```json
/// {
///   "sheet_password": "sheet_secret",
///   "hide_columns_qv": true,
///   "hide_language_sheet": false,
///   "workbook_password": "wb_secret",
///   "row_grouping": {
///     "groups": [{ "start_row": 10, "end_row": 20, "collapsed": false }],
///     "symbols_above": false
///   }
/// }
/// ```
#[derive(Debug, Clone, PartialEq, Default, Builder)]
#[builder(default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportOptions {
    /// Sheet protection password (`None` = no protection, `""` = protection without password)
    #[builder(setter(into, strip_option))]
    pub sheet_password: Option<String>,
    /// Hide helper columns Q-V
    pub hide_columns_qv: bool,
    /// Hide language sheet
    pub hide_language_sheet: bool,
    /// Workbook protection password (`None` = no workbook protection)
    #[builder(setter(into, strip_option))]
    pub workbook_password: Option<String>,
    /// Row grouping for collapsible sections
    #[builder(setter(strip_option))]
    #[cfg_attr(feature = "serde", serde(default))]
    pub row_grouping: Option<RowGrouping>,
}

// ============================================================================
// ReportConfig — Top-level compositor
// ============================================================================

/// Unified configuration for a financial report.
///
/// Composes four domain sub-structs: [`ReportHeader`], [`ReportBody`],
/// [`ReportFooter`], and [`ReportOptions`].
///
/// ## Builder Example
///
/// ```ignore
/// use fb_rust::*;
///
/// let config = ReportConfigBuilder::default()
///     .header(
///         ReportHeaderBuilder::default()
///             .language(Language::Deutsch)
///             .currency(Currency::EUR)
///             .project_number("PROJ-001")
///             .build()?
///     )
///     .footer(
///         ReportFooterBuilder::default()
///             .bank(8500.0)
///             .build()?
///     )
///     .options(
///         ReportOptionsBuilder::default()
///             .sheet_password("geheim")
///             .build()?
///     )
///     .build()?;
/// config.write_to("report.xlsx")?;
/// ```
///
/// ## JSON (minimal)
///
/// ```json
/// {}
/// ```
///
/// All sub-structs default when omitted.
///
/// ## JSON (full)
///
/// ```json
/// {
///   "header": {
///     "language": "english",
///     "currency": "USD",
///     "project_number": "PROJ-2025-001",
///     "project_title": "Education Project",
///     "project_start": "01.01.2025",
///     "project_end": "31.12.2025",
///     "report_start": "01.01.2025",
///     "report_end": "30.06.2025"
///   },
///   "body": {
///     "table": [{ "index": 0, "approved_budget": 50000.0, "income_report": 25000.0, "income_total": 25000.0, "reason": "Donation" }],
///     "left_panel": [{ "index": 0, "date": "15.01.2025", "amount_euro": 1000.0, "amount_local": 1100.0 }],
///     "right_panel": [],
///     "positions": [{ "category": 1, "position": 1, "description": "Personnel", "approved": 18000.0 }],
///     "body_positions": { "1": 10, "2": 5, "6": 0 }
///   },
///   "footer": { "bank": 8500.0, "kasse": 250.50 },
///   "options": {
///     "sheet_password": "sheet_secret",
///     "workbook_password": "wb_secret",
///     "hide_columns_qv": true,
///     "hide_language_sheet": true,
///     "row_grouping": {
///       "groups": [{ "start_row": 10, "end_row": 20, "collapsed": false }],
///       "symbols_above": false
///     }
///   }
/// }
/// ```
///
/// ## Errors
///
/// See [`ReportError`](crate::ReportError) for error variants returned by
/// [`write_to`](Self::write_to) and [`write_to_precomputed`](Self::write_to_precomputed).
#[derive(Debug, Clone, PartialEq, Default, Builder)]
#[builder(default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(deny_unknown_fields))]
pub struct ReportConfig {
    /// Header metadata (language, currency, project info, dates)
    #[cfg_attr(feature = "serde", serde(default))]
    pub header: ReportHeader,
    /// Body data (income table, panels, cost positions)
    #[cfg_attr(feature = "serde", serde(default))]
    pub body: ReportBody,
    /// Footer balances (bank, kasse, sonstiges)
    #[cfg_attr(feature = "serde", serde(default))]
    pub footer: ReportFooter,
    /// Output options (protection, visibility, grouping)
    #[cfg_attr(feature = "serde", serde(default))]
    pub options: ReportOptions,
}

impl ReportConfig {
    /// Writes the financial report to the given file path.
    ///
    /// # Errors
    ///
    /// Returns [`ReportError`](crate::ReportError):
    /// - [`Xlsx`](crate::ReportError::Xlsx) — Excel write failure
    /// - [`Protection`](crate::ReportError::Protection) — workbook password injection failed
    /// - [`Io`](crate::ReportError::Io) — file system error
    /// - [`InvalidPath`](crate::ReportError::InvalidPath) — path contains non-UTF-8
    pub fn write_to(&self, output_path: impl AsRef<Path>) -> Result<(), crate::error::ReportError> {
        self.validate()?;
        let values = self.build_values();
        let body_config = self.build_body_config();
        let sheet_opts = self.build_sheet_options();
        let wb_prot = self
            .options
            .workbook_password
            .as_ref()
            .map(WorkbookProtection::new);
        create_protected_report(
            output_path,
            &values,
            &body_config,
            &sheet_opts,
            wb_prot.as_ref(),
            self.options.hide_language_sheet,
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
        self.validate()?;
        let values = self.build_values();
        let body_config = self.build_body_config();
        let sheet_opts = self.build_sheet_options();
        create_protected_report_precomputed(
            output_path,
            &values,
            &body_config,
            &sheet_opts,
            self.options.hide_language_sheet,
            hash,
        )?;
        Ok(())
    }

    fn validate(&self) -> Result<(), crate::error::ReportError> {
        use std::collections::HashSet;

        // Language and Currency are validated at the type level (Language enum, Currency struct).

        // Table index: 0-4
        for e in &self.body.table {
            if e.index > 4 {
                return Err(crate::error::ReportError::Validation(format!(
                    "table entry index {} out of range (must be 0-4)",
                    e.index
                )));
            }
        }

        // Panel index: 0-17
        for (panel_name, entries) in [
            ("left_panel", &self.body.left_panel),
            ("right_panel", &self.body.right_panel),
        ] {
            for e in entries {
                if e.index > 17 {
                    return Err(crate::error::ReportError::Validation(format!(
                        "{panel_name} entry index {} out of range (must be 0-17)",
                        e.index
                    )));
                }
            }
        }

        // Positions: category 1-8, position vs body_positions, duplicates
        let mut seen = HashSet::new();
        for e in &self.body.positions {
            if e.category < 1 || e.category > 8 {
                return Err(crate::error::ReportError::Validation(format!(
                    "position category {} out of range (must be 1-8)",
                    e.category
                )));
            }

            if !seen.insert((e.category, e.position)) {
                return Err(crate::error::ReportError::Validation(format!(
                    "duplicate position: category {}, position {}",
                    e.category, e.position
                )));
            }

            let max_positions = self
                .body
                .body_positions
                .get(&e.category)
                .copied()
                .unwrap_or_else(|| {
                    BodyConfig::default_positions()
                        .get(&e.category)
                        .copied()
                        .unwrap_or(0)
                });

            if max_positions == 0 {
                // Header-input mode: only position 0 allowed
                if e.position != 0 {
                    return Err(crate::error::ReportError::Validation(format!(
                        "category {} has 0 positions (header-input mode), but position {} was given (must be 0)",
                        e.category, e.position
                    )));
                }
            } else {
                // Normal mode: position must be 1..=max_positions
                if e.position == 0 || e.position > max_positions {
                    return Err(crate::error::ReportError::Validation(format!(
                        "category {} has {} positions, but position {} was given (must be 1-{})",
                        e.category, max_positions, e.position, max_positions
                    )));
                }
            }
        }

        Ok(())
    }

    fn build_values(&self) -> ReportValues {
        let h = &self.header;
        let mut v = ReportValues::new()
            .with_lang(h.language)
            .with_curr(h.currency);

        if let Some(ref s) = h.project_number {
            v = v.with_project_number(s);
        }
        if let Some(ref s) = h.project_title {
            v = v.with_project_title(s);
        }
        if let Some(ref s) = h.project_start {
            v = v.with_project_start(s);
        }
        if let Some(ref s) = h.project_end {
            v = v.with_project_end(s);
        }
        if let Some(ref s) = h.report_start {
            v = v.with_report_start(s);
        }
        if let Some(ref s) = h.report_end {
            v = v.with_report_end(s);
        }

        // Table (rows 15-19)
        for e in &self.body.table {
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

        // Left panel
        for e in &self.body.left_panel {
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

        // Right panel
        for e in &self.body.right_panel {
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

        // Cost positions
        for e in &self.body.positions {
            if e.position == 0 {
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

        // Footer balances
        if let Some(n) = self.footer.bank {
            v.set_footer_bank(n);
        }
        if let Some(n) = self.footer.kasse {
            v.set_footer_kasse(n);
        }
        if let Some(n) = self.footer.sonstiges {
            v.set_footer_sonstiges(n);
        }

        v
    }

    fn build_body_config(&self) -> BodyConfig {
        let mut config = BodyConfig::new();
        for (&cat, &count) in &self.body.body_positions {
            config = config.with_positions(cat, count);
        }
        config
    }

    fn build_sheet_options(&self) -> SheetOptions {
        let mut opts = if let Some(ref pw) = self.options.sheet_password {
            let mut prot = SheetProtection::from_defaults();
            if !pw.is_empty() {
                prot = prot.with_password(pw);
            }
            SheetOptions::new().with_protection(prot)
        } else {
            SheetOptions::new()
        };
        if self.options.hide_columns_qv {
            opts = opts.with_hidden_columns_qv();
        }
        if let Some(ref rg) = self.options.row_grouping {
            opts = opts.with_row_grouping(rg.clone());
        }
        opts
    }
}

// ============================================================================
// Helper functions
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

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_config() -> ReportConfig {
        ReportConfig {
            body: ReportBody {
                table: vec![TableEntry { index: 0, ..Default::default() }],
                positions: vec![
                    PositionEntry { category: 1, position: 1, ..Default::default() },
                    PositionEntry { category: 6, position: 0, ..Default::default() },
                ],
                body_positions: [(1u8, 5u16), (6, 0)].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        }
    }

    #[test]
    fn test_validate_valid_config() {
        assert!(valid_config().validate().is_ok());
    }

    #[test]
    fn test_validate_default_config() {
        assert!(ReportConfig::default().validate().is_ok());
    }

    #[test]
    fn test_validate_table_index_out_of_range() {
        let config = ReportConfig {
            body: ReportBody {
                table: vec![TableEntry { index: 5, ..Default::default() }],
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("table entry index 5"));
    }

    #[test]
    fn test_validate_left_panel_index_out_of_range() {
        let config = ReportConfig {
            body: ReportBody {
                left_panel: vec![PanelEntry { index: 18, ..Default::default() }],
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("left_panel entry index 18"));
    }

    #[test]
    fn test_validate_right_panel_index_out_of_range() {
        let config = ReportConfig {
            body: ReportBody {
                right_panel: vec![PanelEntry { index: 20, ..Default::default() }],
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("right_panel entry index 20"));
    }

    #[test]
    fn test_validate_category_out_of_range_zero() {
        let config = ReportConfig {
            body: ReportBody {
                positions: vec![PositionEntry { category: 0, position: 1, ..Default::default() }],
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 0 out of range"));
    }

    #[test]
    fn test_validate_category_out_of_range_nine() {
        let config = ReportConfig {
            body: ReportBody {
                positions: vec![PositionEntry { category: 9, position: 1, ..Default::default() }],
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 9 out of range"));
    }

    #[test]
    fn test_validate_position_exceeds_body_positions() {
        let config = ReportConfig {
            body: ReportBody {
                positions: vec![PositionEntry { category: 1, position: 6, ..Default::default() }],
                body_positions: [(1u8, 5u16)].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 1 has 5 positions, but position 6"));
    }

    #[test]
    fn test_validate_position_zero_for_normal_category() {
        let config = ReportConfig {
            body: ReportBody {
                positions: vec![PositionEntry { category: 1, position: 0, ..Default::default() }],
                body_positions: [(1u8, 5u16)].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 1 has 5 positions, but position 0"));
    }

    #[test]
    fn test_validate_position_nonzero_for_header_input() {
        let config = ReportConfig {
            body: ReportBody {
                positions: vec![PositionEntry { category: 6, position: 1, ..Default::default() }],
                body_positions: [(6u8, 0u16)].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("category 6 has 0 positions (header-input mode)"));
    }

    #[test]
    fn test_validate_duplicate_position() {
        let config = ReportConfig {
            body: ReportBody {
                positions: vec![
                    PositionEntry { category: 1, position: 1, ..Default::default() },
                    PositionEntry { category: 1, position: 1, ..Default::default() },
                ],
                body_positions: [(1u8, 5u16)].into_iter().collect(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        let err = config.validate().unwrap_err();
        assert!(err.to_string().contains("duplicate position: category 1, position 1"));
    }

    #[test]
    fn test_validate_uses_default_body_positions() {
        // Category 1 not in body_positions → falls back to default (20)
        let config = ReportConfig {
            body: ReportBody {
                positions: vec![PositionEntry { category: 1, position: 15, ..Default::default() }],
                body_positions: HashMap::new(),
                ..ReportBody::default()
            },
            ..ReportConfig::default()
        };
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_typed_language_in_header() {
        let config = ReportConfig {
            header: ReportHeader {
                language: Language::English,
                ..ReportHeader::default()
            },
            ..ReportConfig::default()
        };
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_typed_currency_in_header() {
        let config = ReportConfig {
            header: ReportHeader {
                currency: Currency::USD,
                ..ReportHeader::default()
            },
            ..ReportConfig::default()
        };
        assert!(config.validate().is_ok());
    }
}
