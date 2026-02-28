//! Report writer with registry-based formula evaluation.
//!
//! Die Haupt-API ist `write_report_with_options()`.

use super::layout::{self, setup_sheet};
use super::structure::write_structure;
use crate::report::api::{CellValue, ReportValues};
use crate::report::body::{
    register_body_formulas, register_footer_formulas, BodyConfig, BodyLayout, FooterLayout,
};
use crate::report::core::{build_registry, CellAddr, CellKind, CellRegistry, EvalContext};

// Type alias for complex CellRegistry type
type DynCellRegistry = CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>;
use crate::report::format::{
    build_format_matrix, extend_format_matrix_with_body, extend_format_matrix_with_footer,
    extend_format_matrix_with_prebody, FormatMatrix, ReportOptions, ReportStyles, SectionStyles,
};
use rust_xlsxwriter::{Format, Formula, Workbook, Worksheet, XlsxError};
use std::collections::HashMap;
use std::path::Path;

/// Ergebnis der Body-Generierung
#[derive(Debug, Clone)]
pub struct BodyResult {
    /// Das berechnete Layout
    pub layout: BodyLayout,
    /// Letzte beschriebene Zeile
    pub last_row: u32,
    /// Zeile der Gesamt-Summe
    pub total_row: u32,
    /// E-Spalte Total (für Footer Check-Formel)
    pub e_total: Option<f64>,
    /// F-Spalte Total (für Footer Check-Formel)
    pub f_total: Option<f64>,
}

/// Interne Funktion: Schreibt den Report mit Body-Bereich
fn write_report_with_body(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
) -> Result<BodyResult, XlsxError> {
    // Standardwert für Formel-Ergebnisse auf "" setzen (statt 0)
    ws.set_formula_result_default("");

    // 1. Registry erstellen (statische Zellen)
    let mut registry = build_registry()
        .map_err(|e| XlsxError::ParameterError(format!("Registry error: {}", e)))?;

    // 2. Layouts berechnen
    let body_layout = BodyLayout::compute(body_config);
    let footer_layout = FooterLayout::compute(body_layout.total_row);
    let income_row = 19u32;

    // 3. Dynamische Registrierung (Body + Footer Formeln + API-Zellen)
    register_body_formulas(&mut registry, &body_layout)
        .map_err(|e| XlsxError::ParameterError(format!("Body registry error: {}", e)))?;
    register_footer_formulas(&mut registry, &footer_layout, income_row)
        .map_err(|e| XlsxError::ParameterError(format!("Footer registry error: {}", e)))?;

    // 4. Alle Zellen evaluieren (statisch + dynamisch, topologisch sortiert)
    let computed = evaluate_all_cells(&registry, values)?;

    // 5. FormatMatrix vollständig aufbauen (statisch + body + footer)
    let sec = SectionStyles::new(styles);
    let mut fmt = build_format_matrix(styles, &sec);
    extend_format_matrix_with_body(&mut fmt, styles, &body_layout);
    extend_format_matrix_with_footer(&mut fmt, styles, &sec, footer_layout.start_row);
    extend_format_matrix_with_prebody(&mut fmt, &sec);

    // 6. Komplette Struktur schreiben (Merges, Blanks, statische Strings)
    let lang = values.language().unwrap_or("");
    write_structure(ws, &fmt, &body_layout, &footer_layout, suffix, lang)?;

    // 7. ALLE Zellen aus Registry schreiben (Formeln + API-Werte)
    write_cells_from_registry(ws, &registry, &computed, &fmt)?;

    // 8. Freeze Pane
    layout::setup_freeze_panes(ws, 9)?;

    // BodyResult aus computed
    let last_row = body_layout.last_row;
    let total_row = body_layout.total_row;
    let e_total = computed
        .get(&CellAddr::new(total_row, 4))
        .and_then(|v| v.as_number());
    let f_total = computed
        .get(&CellAddr::new(total_row, 5))
        .and_then(|v| v.as_number());

    Ok(BodyResult {
        layout: body_layout,
        last_row,
        total_row,
        e_total,
        f_total,
    })
}

/// Schreibt den kompletten Finanzbericht MIT dynamischem Body-Bereich UND Optionen
///
/// Diese Funktion unterstützt:
/// - Sheet Protection mit konfigurierbaren Einstellungen
/// - Data Validation für Input-Felder
/// - Column Hiding (Q:V)
///
/// # Arguments
/// * `ws` - Worksheet
/// * `styles` - Report styles
/// * `suffix` - Dateiname-Suffix
/// * `values` - Eingabewerte
/// * `body_config` - Body-Konfiguration
/// * `options` - Protection, Validation und Display-Optionen
pub fn write_report_with_options(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &ReportOptions,
) -> Result<BodyResult, XlsxError> {
    // Basis-Report schreiben
    let body_result = write_report_with_body(ws, styles, suffix, values, body_config)?;

    // Optionen anwenden
    apply_report_options(ws, options, &body_result)?;

    Ok(body_result)
}

/// Wendet ReportOptions auf ein Worksheet an
///
/// Kann nach dem Schreiben des Reports aufgerufen werden.
pub fn apply_report_options(
    ws: &mut Worksheet,
    options: &ReportOptions,
    _body_result: &BodyResult,
) -> Result<(), XlsxError> {
    // 1. Spalten und Zeilen verstecken
    apply_hidden_ranges(ws, &options.hidden)?;

    // 2. Row Grouping anwenden (wenn konfiguriert)
    apply_row_grouping(ws, &options.row_grouping)?;

    // 3. Sheet Protection anwenden (wenn konfiguriert)
    if let Some(ref protection) = options.protection {
        let prot_options = protection.to_protection_options();

        if let Some(ref password) = protection.password {
            ws.protect_with_password(password);
        }
        ws.protect_with_options(&prot_options);
    }

    Ok(())
}

// ============================================================================
// High-Level API: create_protected_report
// ============================================================================

/// Extracts the language suffix from ReportValues for sheet naming
fn extract_suffix_from_values(values: &ReportValues) -> String {
    use crate::report::api::ApiKey;
    let lang_value = values.get(ApiKey::Language);

    if let Some(lang) = lang_value.as_text() {
        if let Some(suffix) = crate::lang::LANG_SUFFIXES.get(lang) {
            return suffix.to_string();
        }
    }

    String::new()
}

/// Creates a complete financial report with optional workbook protection
///
/// This high-level function orchestrates the entire report creation workflow:
/// 1. Creates Workbook and adds language sheet (unless hidden)
/// 2. Sets up worksheet with correct formatting
/// 3. Writes report content via `write_report_with_options()`
/// 4. Optionally applies workbook-level protection (structure lock)
/// 5. Saves to specified output path
///
/// # Arguments
/// * `output_path` - Path where the final Excel file will be saved
/// * `styles` - Report styling configuration
/// * `values` - Report data (all input cells)
/// * `body_config` - Body layout configuration (categories, positions)
/// * `options` - Report options (protection, validation, hidden ranges, etc.)
///
/// # Returns
/// * `Ok(())` on success
/// * `Err` if any step fails (workbook creation, writing, protection, saving)
///
/// # Example
/// ```ignore
/// use kmw_fb_rust::{
///     create_protected_report, ReportStyles, ReportValues,
///     BodyConfig, ReportOptions,
/// };
///
/// let styles = ReportStyles::new();
/// let mut values = ReportValues::new();
/// values.set_language("deutsch");
/// values.set_currency("EUR");
///
/// let body_config = BodyConfig::default();
/// let options = ReportOptions::with_default_protection()
///     .with_workbook_protection("secret123")
///     .with_hidden_columns_qv();
///
/// create_protected_report(
///     "output/report.xlsx",
///     &styles,
///     &values,
///     &body_config,
///     &options,
/// )?;
/// ```
pub fn create_protected_report(
    output_path: impl AsRef<Path>,
    styles: &ReportStyles,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &ReportOptions,
) -> Result<(), Box<dyn std::error::Error>> {
    let output_path = output_path.as_ref();

    // 1. Create workbook
    let mut workbook = Workbook::new();

    // 2. Create main report worksheet FIRST (so it's the leftmost sheet)
    let ws = workbook.add_worksheet();

    // Set worksheet name based on language
    // LANG_CONFIG uses capitalized keys (e.g., "Deutsch"), but values might be lowercase (e.g., "deutsch")
    if let Some(lang_text) = values.get(crate::report::api::ApiKey::Language).as_text() {
        // Try exact match first
        let config = crate::lang::LANG_CONFIG
            .get(lang_text)
            // If not found, try to find by lang_val (case-insensitive)
            .or_else(|| {
                crate::lang::LANG_CONFIG
                    .values()
                    .find(|c| c.lang_val.eq_ignore_ascii_case(lang_text))
            });

        if let Some(config) = config {
            ws.set_name(config.fb_sheet)?;
        }
    }

    setup_sheet(ws)?;

    // 3. Extract suffix for sheet naming
    let suffix = extract_suffix_from_values(values);

    // 4. Write report content
    write_report_with_options(ws, styles, &suffix, values, body_config, options)?;

    // 5. Add language sheet AFTER main report (rightmost position)
    // Hidden by default if hide_language_sheet is true
    crate::lang::build_sheet_with_visibility(&mut workbook, options.hide_language_sheet)?;

    // 6. Apply workbook protection if configured
    if let Some(wb_prot) = &options.workbook_protection {
        // Create temporary file
        let temp_path = output_path.with_extension("tmp.xlsx");

        // Save unprotected version to temp
        workbook.save(&temp_path)?;

        // Apply workbook protection
        crate::workbook_protection::protect_workbook_with_spin_count(
            temp_path.to_str().ok_or("Invalid temp path")?,
            output_path.to_str().ok_or("Invalid output path")?,
            &wb_prot.password,
            wb_prot.spin_count,
        )?;

        // Clean up temp file
        std::fs::remove_file(temp_path)?;
    } else {
        // Save directly without protection
        workbook.save(output_path)?;
    }

    Ok(())
}

/// Wendet HiddenRanges auf ein Worksheet an
fn apply_hidden_ranges(
    ws: &mut Worksheet,
    hidden: &crate::report::format::HiddenRanges,
) -> Result<(), XlsxError> {
    // Spalten verstecken
    for range in hidden.column_ranges() {
        for col in range.start..=range.end {
            ws.set_column_hidden(col as u16)?;
        }
    }

    // Zeilen verstecken
    for range in hidden.row_ranges() {
        for row in range.start..=range.end {
            ws.set_row_hidden(row)?;
        }
    }

    Ok(())
}

/// Wendet Row Grouping auf ein Worksheet an
fn apply_row_grouping(
    ws: &mut Worksheet,
    grouping: &crate::report::format::RowGrouping,
) -> Result<(), XlsxError> {
    if grouping.is_empty() {
        return Ok(());
    }

    // Outline-Symbole Position setzen
    if grouping.symbols_above {
        ws.group_symbols_above(true);
    }

    // Gruppen anwenden
    for group in grouping.groups() {
        if group.collapsed {
            ws.group_rows_collapsed(group.start_row, group.end_row)?;
        } else {
            ws.group_rows(group.start_row, group.end_row)?;
        }
    }

    Ok(())
}

/// Evaluates all cells and returns computed values.
///
/// Verwendet topologische Sortierung (Kahn's Algorithmus) um sicherzustellen,
/// dass Formel-Dependencies vor ihren Abhängigen evaluiert werden.
fn evaluate_all_cells(
    registry: &DynCellRegistry,
    values: &ReportValues,
) -> Result<HashMap<CellAddr, CellValue>, XlsxError> {
    let mut computed: HashMap<CellAddr, CellValue> = HashMap::new();

    // 1. API-Werte eintragen
    for addr in registry.api_cells() {
        if let Some(CellKind::Api(api)) = registry.get(*addr) {
            let value = get_api_value(values, api.key);
            computed.insert(*addr, value);
        }
    }

    // 2. Formeln evaluieren (in topologischer Reihenfolge)
    let eval_order = registry
        .get_eval_order()
        .map_err(|e| XlsxError::ParameterError(format!("Eval order error: {}", e)))?;

    for addr in eval_order {
        if let Some(CellKind::Formula(f)) = registry.get(addr) {
            let ctx = EvalContext {
                computed: &computed,
                api_values: values,
            };
            let result = (f.eval)(&ctx);
            computed.insert(addr, result);
        }
    }

    Ok(computed)
}

/// Holt API-Wert aus ReportValues
///
/// Verwendet `get_owned()` um alle Keys inkl. Footer-Keys zu unterstützen.
fn get_api_value(values: &ReportValues, key: crate::report::api::ApiKey) -> CellValue {
    values.get_owned(key)
}

/// Schreibt alle Zellen aus der Registry (API-Werte und Formeln mit Cache)
fn write_cells_from_registry(
    ws: &mut Worksheet,
    registry: &DynCellRegistry,
    computed: &HashMap<CellAddr, CellValue>,
    fmt: &FormatMatrix,
) -> Result<(), XlsxError> {
    // 1. API-Zellen schreiben (mit ihren berechneten Werten, auch leere als Blanks)
    for addr in registry.api_cells() {
        if let Some(value) = computed.get(addr) {
            write_cell_value(ws, *addr, value, fmt)?;
        }
    }

    // 2. Formel-Zellen schreiben (mit gecachten Ergebnissen)
    for addr in registry.formula_cells() {
        if let Some(CellKind::Formula(f)) = registry.get(*addr) {
            let result = computed.get(addr).cloned().unwrap_or(CellValue::Empty);

            // Formula mit Result erstellen (Cache für Excel)
            let formula = Formula::new(f.excel).set_result(cell_value_to_string(&result));

            // Mit Format schreiben (locked)
            if let Some(format) = fmt.get_locked(addr.row, addr.col) {
                ws.write_formula_with_format(addr.row, addr.col, formula, format)?;
            } else {
                let locked = Format::new().set_locked();
                ws.write_formula_with_format(addr.row, addr.col, formula, &locked)?;
            }
        }
    }

    Ok(())
}

/// Konvertiert CellValue zu String für Formula::set_result
fn cell_value_to_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(s) => s.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Date(d) => d.clone(),
    }
}

/// Schreibt einen Zellwert (für Registry-API-Cells)
fn write_cell_value(
    ws: &mut Worksheet,
    addr: CellAddr,
    value: &CellValue,
    fmt: &FormatMatrix,
) -> Result<(), XlsxError> {
    let format = fmt.get(addr.row, addr.col);

    match value {
        CellValue::Empty => {
            if let Some(f) = format {
                ws.write_blank(addr.row, addr.col, f)?;
            }
        }
        CellValue::Text(s) => {
            if let Some(f) = format {
                ws.write_string_with_format(addr.row, addr.col, s, f)?;
            } else {
                ws.write_string(addr.row, addr.col, s)?;
            }
        }
        CellValue::Number(n) => {
            if let Some(f) = format {
                ws.write_number_with_format(addr.row, addr.col, *n, f)?;
            } else {
                ws.write_number(addr.row, addr.col, *n)?;
            }
        }
        CellValue::Date(d) => {
            if let Some(f) = format {
                ws.write_string_with_format(addr.row, addr.col, d, f)?;
            } else {
                ws.write_string(addr.row, addr.col, d)?;
            }
        }
    }

    Ok(())
}
