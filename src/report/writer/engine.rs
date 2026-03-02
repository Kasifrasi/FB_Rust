//! Report writer with IronCalc-based formula evaluation.
//!
//! Die Haupt-API ist `write_report_with_options()`.

use super::layout::{self, setup_sheet};
use super::structure::write_structure;
use crate::report::api::{CellValue, ReportValues};
use crate::report::body::config::BODY_START_ROW;
use crate::report::body::{BodyConfig, BodyLayout, CategoryMode, FooterLayout};
use crate::report::calc::{CalcBridge, ModelTemplate};
use crate::report::core::CellAddr;
use crate::report::options::SheetOptions;
use crate::report::styles::{
    build_format_matrix, extend_format_matrix_with_body, extend_format_matrix_with_footer,
    extend_format_matrix_with_prebody, FormatMatrix, ReportStyles, SectionStyles,
};
use rust_xlsxwriter::{Format, Formula, Workbook, Worksheet, XlsxError};
use std::path::Path;
use std::sync::LazyLock;

/// Globales Master-Template: Einmal erstellt, für alle Reports wiederverwendet.
///
/// Enthält alle statischen Formeln + Sprachversionen-Daten.
/// Wird beim ersten Zugriff initialisiert (LazyLock).
static MASTER_TEMPLATE: LazyLock<ModelTemplate> = LazyLock::new(ModelTemplate::new);

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

    // 1. CalcBridge aus Template erstellen (enthält statische Formeln + Sprachversionen)
    let mut bridge = CalcBridge::from_template(&MASTER_TEMPLATE);

    // 2. Layouts berechnen
    let body_layout = BodyLayout::compute(body_config);
    let footer_layout = FooterLayout::compute(body_layout.total_row);
    let income_row = 19u32;

    // 3. Dynamische Formeln registrieren (Body + Footer)
    bridge.register_body_formulas(&body_layout);
    bridge.register_footer_formulas(&footer_layout, income_row);

    // 4. Input-Werte setzen + Evaluation
    bridge.populate(values, &body_layout, &footer_layout);
    bridge.evaluate();

    // 5. FormatMatrix vollständig aufbauen (statisch + body + footer)
    let sec = SectionStyles::new(&styles);
    let mut fmt = build_format_matrix(&styles, &sec);
    extend_format_matrix_with_body(&mut fmt, &styles, &body_layout);
    extend_format_matrix_with_footer(&mut fmt, &styles, &sec, footer_layout.start_row);
    extend_format_matrix_with_prebody(&mut fmt, &sec);

    // 6. Komplette Struktur schreiben (Merges, Blanks, statische Strings)
    let lang = values.language().unwrap_or("");
    write_structure(ws, &fmt, &body_layout, &footer_layout, suffix, lang)?;

    // 7. ALLE Zellen aus Bridge schreiben (Formeln + Input-Werte)
    write_cells_from_bridge(ws, &bridge, &fmt)?;

    // 8. Freeze Pane
    layout::setup_freeze_panes(ws, 9)?;

    // BodyResult
    let last_row = body_layout.last_row;
    let total_row = body_layout.total_row;
    let e_total = bridge.get_value(total_row, 4).as_number();
    let f_total = bridge.get_value(total_row, 5).as_number();

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
            crate::lang::LANG_CONFIG.get(lang_text).or_else(|| {
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
fn compute_page_breaks(layout: &BodyLayout, footer_layout: &FooterLayout) -> Vec<u32> {
    let mut breaks = Vec::new();
    let mut rows_on_page = BODY_START_ROW;

    for cat in &layout.categories {
        let is_cat_8 = cat.meta.num == 8;

        let cat_start = match &cat.mode {
            CategoryMode::HeaderInput { row } => *row,
            CategoryMode::WithPositions { header_row, .. } => *header_row,
        };
        let mut block_end = cat.sum_row();
        if is_cat_8 {
            // Kategorie 8 wird immer mit der Summenzeile verbunden
            block_end = layout.total_row;
        }

        let block_rows = block_end - cat_start + 1;

        if rows_on_page + block_rows <= MAX_ROWS_PER_PAGE {
            // Passt noch auf die aktuelle Seite
            rows_on_page += block_rows;
        } else {
            // Passt NICHT mehr auf die aktuelle Seite.
            // Regel: Eine Kategorie wird nur dann auf der nächsten Seite komplett gehalten,
            // wenn die aktuelle Seite dadurch mindestens ~65 Zeilen (2/3 von MAX_ROWS_PER_PAGE) nutzt.
            // Ansonsten füllen wir die Seite auf und brechen innerhalb der Kategorie um.
            // Kategorie 8 ist eine Ausnahme und wird immer als Block auf die nächste Seite verschoben
            // (oder ggf. dort umbrochen, falls sie allein zu groß ist).
            let min_rows_to_keep = 65;
            let should_fill_page =
                !is_cat_8 && rows_on_page < min_rows_to_keep && cat_start > BODY_START_ROW;

            if should_fill_page {
                // Seite auffüllen -> harter Umbruch innerhalb der Kategorie
                let page_remaining = MAX_ROWS_PER_PAGE.saturating_sub(rows_on_page);
                let mut pos = cat_start + page_remaining;
                while pos <= block_end {
                    breaks.push(pos);
                    pos += MAX_ROWS_PER_PAGE;
                }
                let last_break = breaks.last().copied().unwrap_or(cat_start);
                rows_on_page = block_end.saturating_sub(last_break) + 1;
            } else {
                // Umbruch VOR dem Block, sofern wir nicht eh am Anfang sind
                if cat_start > BODY_START_ROW {
                    breaks.push(cat_start);
                    rows_on_page = 0;
                }

                // Prüfen ob der Block jetzt auf die leere Seite passt
                if rows_on_page + block_rows > MAX_ROWS_PER_PAGE {
                    // Block ist allein schon zu groß, muss hart umbrochen werden
                    let start_of_page = if rows_on_page == 0 {
                        cat_start
                    } else {
                        cat_start.saturating_sub(rows_on_page)
                    };
                    let mut pos = start_of_page + MAX_ROWS_PER_PAGE;
                    while pos <= block_end {
                        breaks.push(pos);
                        pos += MAX_ROWS_PER_PAGE;
                    }
                    let last_break = breaks.last().copied().unwrap_or(cat_start);
                    rows_on_page = block_end.saturating_sub(last_break) + 1;
                } else {
                    rows_on_page += block_rows;
                }
            }
        }
    }

    // Footer Block
    let footer_start = footer_layout.start_row;
    let footer_rows = footer_layout.end_row - footer_layout.start_row + 1;

    if rows_on_page + footer_rows > MAX_ROWS_PER_PAGE {
        // Footer passt nicht auf diese Seite -> Umbruch VOR Footer
        breaks.push(footer_start);
    }

    breaks
}

/// Wendet Druckeinstellungen auf das Worksheet an (A4, Hochformat, Seitenumbrüche).
fn apply_print_setup(
    ws: &mut Worksheet,
    body_result: &BodyResult,
    footer_layout: &FooterLayout,
) -> Result<(), XlsxError> {
    ws.set_paper_size(9); // A4
    ws.set_print_fit_to_pages(1, 0); // Breite = 1 Seite, Höhe unbegrenzt

    let breaks = compute_page_breaks(&body_result.layout, footer_layout);
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
    let range = format!("{}!$B$1:$H${},{}!$J$1:$V$31", q, footer_excel, q);
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
    apply_print_setup(ws, &body_result, &footer_layout)?;

    setup_print_area(&mut workbook, &sheet_name, footer_layout.end_row)?;
    crate::lang::build_sheet_with_visibility(&mut workbook, hide_language_sheet)?;

    if let Some(wb_prot) = wb_protection {
        let tmp = tempfile::NamedTempFile::new_in(output_path.parent().unwrap_or(Path::new(".")))?;
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
    apply_print_setup(ws, &body_result, &footer_layout)?;

    setup_print_area(&mut workbook, &sheet_name, footer_layout.end_row)?;
    crate::lang::build_sheet_with_visibility(&mut workbook, hide_language_sheet)?;

    let tmp = tempfile::NamedTempFile::new_in(output_path.parent().unwrap_or(Path::new(".")))?;
    workbook.save(tmp.path())?;

    crate::workbook_protection::protect_workbook_precomputed(
        tmp.path()
            .to_str()
            .ok_or_else(|| crate::error::ReportError::InvalidPath(format!("{:?}", tmp.path())))?,
        output_path
            .to_str()
            .ok_or_else(|| crate::error::ReportError::InvalidPath(format!("{:?}", output_path)))?,
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

/// Schreibt alle Zellen aus der CalcBridge (Input-Werte und Formeln mit Cache)
fn write_cells_from_bridge(
    ws: &mut Worksheet,
    bridge: &CalcBridge,
    fmt: &FormatMatrix,
) -> Result<(), XlsxError> {
    // 1. Input-Zellen schreiben (mit ihren Werten, auch leere als Blanks)
    for addr in bridge.input_cells() {
        let value = bridge.get_value(addr.row, addr.col);
        write_cell_value(ws, *addr, &value, fmt)?;
    }

    // 2. Formel-Zellen schreiben (mit gecachten Ergebnissen)
    for addr in bridge.formula_cells() {
        if let Some(formula_str) = bridge.get_formula(addr.row, addr.col) {
            let result = bridge.get_value(addr.row, addr.col);

            // Formula mit Result erstellen (Cache für Excel)
            let formula = Formula::new(&formula_str).set_result(cell_value_to_string(&result));

            // Mit Format schreiben (locked)
            if let Some(format) = fmt.get_locked(addr.row, addr.col) {
                ws.write_formula_with_format(addr.row, addr.col, formula, format)?;
            } else {
                let locked = Format::new().set_locked();
                ws.write_formula_with_format(addr.row, addr.col, formula, &locked)?;
            }
        }
    }

    // 3. Hyperlink-Zellen: Excel-Formel mit IronCalc-evaluiertem URL-Cache
    //    IronCalc unterstützt HYPERLINK() nicht, daher: VLOOKUP → URL evaluiert,
    //    HYPERLINK(VLOOKUP(...)) als Excel-Formel mit URL als Cache geschrieben.
    for (addr, excel_formula) in bridge.hyperlink_cells() {
        let url_value = bridge.get_value(addr.row, addr.col);
        let formula = Formula::new(excel_formula).set_result(cell_value_to_string(&url_value));

        if let Some(format) = fmt.get_locked(addr.row, addr.col) {
            ws.write_formula_with_format(addr.row, addr.col, formula, format)?;
        } else {
            let locked = Format::new().set_locked();
            ws.write_formula_with_format(addr.row, addr.col, formula, &locked)?;
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
