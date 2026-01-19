//! Writer - Hauptkoordinator für den Finanzbericht
//!
//! Koordiniert das Schreiben aller Sections:
//! - Header Section
//! - Table Section
//! - Panel Section
//! - Formeln (zentral aus HEADER_FORMULAS)

use crate::v2::report::formats::{build_format_matrix, FormatMatrix, ReportStyles, SectionStyles};
use crate::v2::report::formulas::{evaluate_formula, FormulaCache, HEADER_FORMULAS};
use crate::v2::report::layout;
use crate::v2::report::sections::{write_header_section, write_panel_section, write_table_section};
use crate::v2::report::values::{CellValue, ReportValues};
use rust_xlsxwriter::{Formula, Worksheet, XlsxError};

/// Schreibt den kompletten Finanzbericht
pub fn write_report(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    values: &ReportValues,
) -> Result<(), XlsxError> {
    // Styles und FormatMatrix erstellen
    let sec = SectionStyles::new(styles);
    let fmt = build_format_matrix(styles, &sec);

    // Hole lang_val aus ReportValues
    let lang_val = values.language().unwrap_or("");

    // 1. Header Section (Row 0-9)
    write_header_section(ws, &fmt, suffix, lang_val)?;

    // 2. Table Section (Row 10-19)
    write_table_section(ws, &fmt)?;

    // 3. Panel Section (Row 10-30, J-V)
    write_panel_section(ws, &fmt, values)?;

    // 4. Formeln (zentral aus HEADER_FORMULAS)
    write_formulas(ws, &fmt, values)?;

    // 5. Freeze Pane
    layout::setup_freeze_panes(ws, 9)?;

    Ok(())
}

/// Schreibt alle Formeln aus HEADER_FORMULAS mit ausgewerteten Ergebnissen
fn write_formulas(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    values: &ReportValues,
) -> Result<(), XlsxError> {
    let mut cache = FormulaCache::new();

    for formula_def in HEADER_FORMULAS.iter() {
        let row = formula_def.address.row;
        let col = formula_def.address.col;

        // Formel auswerten
        let result = evaluate_formula(formula_def, values, &mut cache);

        // Formula mit Result erstellen
        let formula_with_result =
            Formula::new(&formula_def.excel_formula).set_result(cell_value_to_string(&result));

        // Formel-Zellen werden mit locked Format geschrieben
        if let Some(format) = fmt.get_locked(row, col) {
            ws.write_formula_with_format(row, col, formula_with_result, &format)?;
        } else {
            // Fallback: locked Format ohne spezielle Formatierung
            let locked_format = rust_xlsxwriter::Format::new().set_locked();
            ws.write_formula_with_format(row, col, formula_with_result, &locked_format)?;
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
