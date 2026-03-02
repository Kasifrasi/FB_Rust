//! Report writer with registry-based formula evaluation.
//!
//! Die Haupt-API ist `write_report_with_options()`.

use super::layout::{self, setup_sheet};
use super::structure::write_structure;
use crate::report::api::{CellValue, ReportValues};
use crate::report::body::{
    register_body_formulas, register_footer_formulas, BodyConfig, BodyLayout, CategoryMode,
    FooterLayout,
};
use crate::report::body::config::BODY_START_ROW;
use crate::report::core::{build_registry, CellAddr, CellKind, DynRegistry, EvalContext};
use crate::report::options::SheetOptions;
use crate::report::styles::{
    build_format_matrix, extend_format_matrix_with_body, extend_format_matrix_with_footer,
    extend_format_matrix_with_prebody, FormatMatrix, ReportStyles, SectionStyles,
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
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
) -> Result<BodyResult, XlsxError> {
    let styles = ReportStyles::new();
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
    let sec = SectionStyles::new(&styles);
    let mut fmt = build_format_matrix(&styles, &sec);
    extend_format_matrix_with_body(&mut fmt, &styles, &body_layout);
    extend_format_matrix_with_footer(&mut fmt, &styles, &sec, footer_layout.start_row);
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
/// Styles werden intern mit Standard-Einstellungen erzeugt.
///
/// # Arguments
/// * `ws` - Worksheet
/// * `suffix` - Dateiname-Suffix
/// * `values` - Eingabewerte
/// * `body_config` - Body-Konfiguration
/// * `options` - Protection, Validation und Display-Optionen
pub fn write_report_with_options(
    ws: &mut Worksheet,
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &SheetOptions,
) -> Result<BodyResult, XlsxError> {
    // Basis-Report schreiben
    let body_result = write_report_with_body(ws, suffix, values, body_config)?;

    // Optionen anwenden
    apply_sheet_options(ws, options, &body_result)?;

    Ok(body_result)
}

/// Wendet SheetOptions auf ein Worksheet an
///
/// Kann nach dem Schreiben des Reports aufgerufen werden.
pub fn apply_sheet_options(
    ws: &mut Worksheet,
    options: &SheetOptions,
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

/// Extracts the version string for cell B2 from ReportValues.
///
/// Returns the version if set, otherwise an empty string.
fn extract_version_from_values(values: &ReportValues) -> String {
    values.version().unwrap_or("").to_string()
}

// ============================================================================
// Print Setup: Druckbereich + Seitenumbrüche
// ============================================================================

/// Geschätzte Zeilen pro Druckseite (A4 Hochformat, fit_to_pages Breite=1).
const MAX_ROWS_PER_PAGE: u32 = 80;

/// Bestimmt den Sheet-Namen anhand der Spracheinstellung.
///
/// Gibt den lokalisierten Sheet-Namen zurück (z.B. "Finanzbericht", "Financial Report")
/// oder "Sheet1" als Fallback.
fn determine_sheet_name(values: &ReportValues) -> String {
    values
        .get(crate::report::api::ApiKey::Language)
        .as_text()
        .and_then(|lang_text| {
            crate::lang::LANG_CONFIG
                .get(lang_text)
                .or_else(|| {
                    crate::lang::LANG_CONFIG
                        .values()
                        .find(|c| c.lang_val.eq_ignore_ascii_case(lang_text))
                })
        })
        .map(|c| c.fb_sheet.to_string())
        .unwrap_or_else(|| "Sheet1".to_string())
}

/// Berechnet horizontale Seitenumbrüche an Kategorie-Grenzen.
///
/// Platziert Umbrüche VOR Kategorie-Headern, sobald die aktuelle Seite voll ist.
/// Die erste Seite hat weniger Platz, da der Header-Abschnitt (Zeilen 0–25) verbraucht wird.
///
/// **Sonderfall übergroße Kategorie:** Wenn eine einzelne Kategorie zusammen mit dem
/// bisherigen Seiteninhalt mehr als `MAX_ROWS_PER_PAGE` Zeilen belegt, wird zusätzlich
/// innerhalb der Kategorie alle `MAX_ROWS_PER_PAGE` Zeilen hart umbrochen.
/// Für nachfolgende Kategorien gilt wieder die normale Regel (Umbruch vor Header).
fn compute_page_breaks(layout: &BodyLayout) -> Vec<u32> {
    let mut breaks = Vec::new();
    let mut rows_on_page = BODY_START_ROW;

    for cat in &layout.categories {
        let cat_start = match &cat.mode {
            CategoryMode::HeaderInput { row } => *row,
            CategoryMode::WithPositions { header_row, .. } => *header_row,
        };
        let cat_end = cat.sum_row();
        let cat_rows = cat_end - cat_start + 1;

        if rows_on_page + cat_rows <= MAX_ROWS_PER_PAGE {
            // Kategorie passt noch auf die aktuelle Seite
            rows_on_page += cat_rows;
        } else if cat_start > BODY_START_ROW && rows_on_page < MAX_ROWS_PER_PAGE {
            // Kategorie passt nicht mehr → Umbruch VOR dem Header
            breaks.push(cat_start);

            // Prüfen ob die Kategorie allein > MAX_ROWS_PER_PAGE ist
            if cat_rows > MAX_ROWS_PER_PAGE {
                // Harte Umbrüche innerhalb der Kategorie
                let mut pos = cat_start + MAX_ROWS_PER_PAGE;
                while pos <= cat_end {
                    breaks.push(pos);
                    pos += MAX_ROWS_PER_PAGE;
                }
                // Restzeilen auf neuer Seite
                let last_break = *breaks.last().unwrap();
                rows_on_page = cat_end - last_break + 1;
            } else {
                rows_on_page = cat_rows;
            }
        } else {
            // Seite bereits voll oder erste Kategorie übergroß → harte Umbrüche
            let page_remaining = MAX_ROWS_PER_PAGE.saturating_sub(rows_on_page);
            let mut pos = cat_start + page_remaining;
            while pos <= cat_end {
                breaks.push(pos);
                pos += MAX_ROWS_PER_PAGE;
            }
            let last_break = breaks.last().copied().unwrap_or(cat_start);
            rows_on_page = cat_end - last_break + 1;
        }
    }

    breaks
}

/// Wendet Druckeinstellungen auf das Worksheet an (A4, Hochformat, Seitenumbrüche).
fn apply_print_setup(ws: &mut Worksheet, body_result: &BodyResult) -> Result<(), XlsxError> {
    ws.set_paper_size(9); // A4
    ws.set_print_fit_to_pages(1, 0); // Breite = 1 Seite, Höhe unbegrenzt

    let breaks = compute_page_breaks(&body_result.layout);
    if !breaks.is_empty() {
        ws.set_page_breaks(&breaks)?;
    }

    Ok(())
}

/// Setzt den Druckbereich als non-contiguous Print Area via `define_name`.
///
/// Bereich 1: B1:H{footer_end} (Hauptbericht — ggf. mehrere Seiten, Hochformat)
/// Bereich 2: J1:V31 (Belegpanel — immer 1 Seite, Hochformat)
fn setup_print_area(
    workbook: &mut Workbook,
    sheet_name: &str,
    footer_end_row: u32,
) -> Result<(), XlsxError> {
    let footer_excel = footer_end_row + 1; // 0-indexed → 1-indexed
    let q = format!("'{}'", sheet_name);

    let name = format!("{}!_xlnm.Print_Area", q);
    let range = format!(
        "{}!$B$1:$H${},{}!$J$1:$V$31",
        q, footer_excel, q
    );
    workbook.define_name(name, &range)?;
    Ok(())
}

// ============================================================================
// High-Level API: create_protected_report
// ============================================================================

/// Erstellt einen kompletten Finanzbericht mit optionalem Workbook-Schutz.
///
/// Styles werden intern erzeugt. Workbook-Protection und Language-Sheet-Sichtbarkeit
/// werden als separate Parameter übergeben (nicht Teil von SheetOptions).
pub(crate) fn create_protected_report(
    output_path: impl AsRef<Path>,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &SheetOptions,
    wb_protection: Option<&crate::workbook_protection::WorkbookProtection>,
    hide_language_sheet: bool,
) -> Result<(), crate::error::ReportError> {
    let output_path = output_path.as_ref();
    let sheet_name = determine_sheet_name(values);

    let mut workbook = Workbook::new();
    let ws = workbook.add_worksheet();
    ws.set_name(&sheet_name)?;

    setup_sheet(ws)?;
    let suffix = extract_version_from_values(values);
    let body_result = write_report_with_options(ws, &suffix, values, body_config, options)?;
    let footer_layout = FooterLayout::compute(body_result.total_row);
    apply_print_setup(ws, &body_result)?;

    setup_print_area(&mut workbook, &sheet_name, footer_layout.end_row)?;
    crate::lang::build_sheet_with_visibility(&mut workbook, hide_language_sheet)?;

    if let Some(wb_prot) = wb_protection {
        let tmp = tempfile::NamedTempFile::new_in(
            output_path.parent().unwrap_or(Path::new(".")),
        )?;
        workbook.save(tmp.path())?;

        crate::workbook_protection::protect_workbook_with_spin_count(
            tmp.path().to_str().ok_or_else(|| {
                crate::error::ReportError::InvalidPath(format!("{:?}", tmp.path()))
            })?,
            output_path.to_str().ok_or_else(|| {
                crate::error::ReportError::InvalidPath(format!("{:?}", output_path))
            })?,
            &wb_prot.password,
            wb_prot.spin_count,
        )?;
        // tmp wird beim Drop automatisch gelöscht — nicht persisten,
        // da output_path bereits die geschützte Datei enthält.
    } else {
        workbook.save(output_path)?;
    }

    Ok(())
}

/// Wie `create_protected_report`, aber mit vorberechnetem Hash (kein SHA-512-Aufwand).
///
/// Für Batch-Operationen: Hash einmal mit `precompute_hash()` berechnen,
/// dann für N Dateien wiederverwenden.
pub(crate) fn create_protected_report_precomputed(
    output_path: impl AsRef<Path>,
    values: &ReportValues,
    body_config: &BodyConfig,
    options: &SheetOptions,
    hide_language_sheet: bool,
    hash: &crate::workbook_protection::PrecomputedHash,
) -> Result<(), crate::error::ReportError> {
    let output_path = output_path.as_ref();
    let sheet_name = determine_sheet_name(values);

    let mut workbook = Workbook::new();
    let ws = workbook.add_worksheet();
    ws.set_name(&sheet_name)?;

    setup_sheet(ws)?;
    let suffix = extract_version_from_values(values);
    let body_result = write_report_with_options(ws, &suffix, values, body_config, options)?;
    let footer_layout = FooterLayout::compute(body_result.total_row);
    apply_print_setup(ws, &body_result)?;

    setup_print_area(&mut workbook, &sheet_name, footer_layout.end_row)?;
    crate::lang::build_sheet_with_visibility(&mut workbook, hide_language_sheet)?;

    let tmp = tempfile::NamedTempFile::new_in(
        output_path.parent().unwrap_or(Path::new(".")),
    )?;
    workbook.save(tmp.path())?;

    crate::workbook_protection::protect_workbook_precomputed(
        tmp.path().to_str().ok_or_else(|| {
            crate::error::ReportError::InvalidPath(format!("{:?}", tmp.path()))
        })?,
        output_path.to_str().ok_or_else(|| {
            crate::error::ReportError::InvalidPath(format!("{:?}", output_path))
        })?,
        hash,
    )?;
    // tmp wird beim Drop automatisch gelöscht — nicht persistieren,
    // da output_path bereits die geschützte Datei enthält.
    Ok(())
}

/// Wendet HiddenRanges auf ein Worksheet an
fn apply_hidden_ranges(
    ws: &mut Worksheet,
    hidden: &crate::report::options::HiddenRanges,
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
    grouping: &crate::report::options::RowGrouping,
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
    registry: &DynRegistry,
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
    registry: &DynRegistry,
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
