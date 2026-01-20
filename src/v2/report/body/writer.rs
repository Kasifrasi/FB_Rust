//! Body Writer - Schreibt die Struktur des dynamischen Body-Bereichs
//!
//! Schreibt:
//! - Kategorie-Header (Nummer + VLOOKUP Label)
//! - Position-Zeilen (Nummer + API-Werte)
//! - Kategorie-Footer (VLOOKUP Sum-Label + SUMPRODUCT Formeln)
//! - Single-Row Kategorien
//! - Gesamt-Zeile (VLOOKUP Label + SUM Formeln)
//! - Ratio-Formeln (G-Spalte)

use super::config::BodyConfig;
use super::layout::{BodyLayout, CategoryLayout, TOTAL_LABEL_INDEX};
use crate::v2::report::api::{ApiKey, PositionField, SingleRowField};
use crate::v2::report::formats::FormatMatrix;
use crate::v2::report::values::{CellValue, ReportValues};
use rust_xlsxwriter::{Format, Formula, Worksheet, XlsxError};

/// Ergebnis der Body-Generierung
#[derive(Debug, Clone)]
pub struct BodyResult {
    /// Das berechnete Layout
    pub layout: BodyLayout,
    /// Letzte beschriebene Zeile
    pub last_row: u32,
    /// Zeile der Gesamt-Summe
    pub total_row: u32,
}

/// Ratio-Formel für G-Spalte
const RATIO_FORMULA: &str = "=IFERROR(INDEX($F$1:$F$1001,ROW())/INDEX($D$1:$D$1001,ROW()),0)";

/// Schreibt die komplette Body-Struktur (ohne API-Werte, nur Blanks)
///
/// Für API-Werte verwende `write_body_structure_with_values()`.
pub fn write_body_structure(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    config: &BodyConfig,
) -> Result<BodyResult, XlsxError> {
    write_body_structure_with_values(ws, fmt, config, None)
}

/// Schreibt die komplette Body-Struktur MIT API-Werten
///
/// # Arguments
/// * `ws` - Worksheet
/// * `fmt` - FormatMatrix
/// * `config` - BodyConfig
/// * `values` - Optional: ReportValues für API-Werte
pub fn write_body_structure_with_values(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    config: &BodyConfig,
    values: Option<&ReportValues>,
) -> Result<BodyResult, XlsxError> {
    // 1. Layout berechnen
    let layout = BodyLayout::compute(config);

    // 2. Kategorien schreiben
    for cat in &layout.categories {
        if cat.is_multi_row() {
            write_multi_row_category(ws, fmt, cat, &layout, values)?;
        } else {
            write_single_row_category(ws, fmt, cat, &layout, values)?;
        }
    }

    // 3. Gesamt-Zeile schreiben
    write_total_row(ws, fmt, &layout)?;

    // 4. Ratio-Formeln anwenden
    write_ratio_formulas(ws, fmt, &layout)?;

    Ok(BodyResult {
        layout: layout.clone(),
        last_row: layout.last_row,
        total_row: layout.total_row,
    })
}

/// Schreibt eine Multi-Row Kategorie (Header + Positionen + Footer)
fn write_multi_row_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
    layout: &BodyLayout,
    values: Option<&ReportValues>,
) -> Result<(), XlsxError> {
    let header_row = cat.header_row.expect("Multi-row must have header");
    let positions = cat
        .positions
        .as_ref()
        .expect("Multi-row must have positions");
    let footer_row = cat.footer_row.expect("Multi-row must have footer");

    // === Header-Zeile ===
    // B: Kategorie-Nummer
    write_with_format(ws, fmt, header_row, 1, &format!("{}.", cat.meta.num))?;

    // C: VLOOKUP Label
    let label_formula = vlookup_formula(cat.meta.label_index);
    write_formula_with_format(ws, fmt, header_row, 2, &label_formula)?;

    // D-H: Blanks
    for col in 3..=7 {
        write_blank_with_format(ws, fmt, header_row, col)?;
    }

    // === Position-Zeilen ===
    for (i, _row) in (positions.start_row..=positions.end_row).enumerate() {
        let pos_num = (i + 1) as u16;

        // B: Positions-Nummer
        write_with_format(
            ws,
            fmt,
            positions.start_row + i as u32,
            1,
            &format!("{}.{}", cat.meta.num, pos_num),
        )?;

        // C-H: API-Werte oder Blanks
        if let Some(values) = values {
            write_position_values(ws, fmt, layout, values, cat.meta.num, pos_num)?;
        } else {
            // Fallback: Blanks
            for col in 2..=7 {
                if col != 6 {
                    // G ist Ratio-Formel
                    write_blank_with_format(ws, fmt, positions.start_row + i as u32, col)?;
                }
            }
        }
    }

    // === Footer-Zeile ===
    // B:C merged: VLOOKUP Sum-Label
    let sum_label_formula = vlookup_formula(cat.meta.sum_label_index);
    write_merged_formula(ws, fmt, footer_row, 1, 2, &sum_label_formula)?;

    // D-F: SUMPRODUCT Formeln
    for col in 3..=5 {
        let formula = sumproduct_formula(col, positions.start_row, positions.end_row);
        write_formula_with_format(ws, fmt, footer_row, col, &formula)?;
    }

    // G: Ratio (wird später geschrieben)
    // H: Blank
    write_blank_with_format(ws, fmt, footer_row, 7)?;

    Ok(())
}

/// Schreibt eine Single-Row Kategorie
fn write_single_row_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
    layout: &BodyLayout,
    values: Option<&ReportValues>,
) -> Result<(), XlsxError> {
    let row = cat.single_row.expect("Single-row must have single_row");

    // B: Kategorie-Nummer
    write_with_format(ws, fmt, row, 1, &format!("{}.", cat.meta.num))?;

    // C: VLOOKUP Label
    let label_formula = vlookup_formula(cat.meta.label_index);
    write_formula_with_format(ws, fmt, row, 2, &label_formula)?;

    // D, E, F, H: API-Werte oder Blanks (G ist Ratio-Formel)
    if let Some(values) = values {
        write_single_row_values(ws, fmt, layout, values, cat.meta.num)?;
    } else {
        // Fallback: Blanks
        for col in [3, 4, 5, 7] {
            // D, E, F, H - nicht G (Ratio)
            write_blank_with_format(ws, fmt, row, col)?;
        }
    }

    Ok(())
}

/// Schreibt die Gesamt-Zeile
fn write_total_row(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
) -> Result<(), XlsxError> {
    let row = layout.total_row;

    // B:C merged: VLOOKUP "Gesamt" Label
    let label_formula = vlookup_formula(TOTAL_LABEL_INDEX);
    write_merged_formula(ws, fmt, row, 1, 2, &label_formula)?;

    // D-F: SUM Formeln (summiert Footer + Single-Rows)
    let footer_rows = layout.footer_rows();
    let single_rows = layout.single_rows();

    for col in 3..=5 {
        let formula = sum_formula(col, &footer_rows, &single_rows);
        write_formula_with_format(ws, fmt, row, col, &formula)?;
    }

    // G: Ratio (wird später geschrieben)
    // H: Blank
    write_blank_with_format(ws, fmt, row, 7)?;

    Ok(())
}

/// Schreibt alle Ratio-Formeln (G-Spalte)
fn write_ratio_formulas(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
) -> Result<(), XlsxError> {
    for row in layout.ratio_rows() {
        write_formula_with_format(ws, fmt, row, 6, RATIO_FORMULA)?;
    }
    Ok(())
}

// ============================================================================
// Formel-Generierung
// ============================================================================

/// Generiert VLOOKUP-Formel: =IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE))
fn vlookup_formula(index: usize) -> String {
    format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
        index
    )
}

/// Generiert SUMPRODUCT-Formel: =SUMPRODUCT(ROUND(D27:D46,2))
fn sumproduct_formula(col: u16, start_row: u32, end_row: u32) -> String {
    let col_letter = col_to_letter(col);
    format!(
        "=SUMPRODUCT(ROUND({}{}:{}{},2))",
        col_letter,
        start_row + 1, // Excel ist 1-basiert
        col_letter,
        end_row + 1
    )
}

/// Generiert SUM-Formel: =SUM(D47+D69+D91+...)
fn sum_formula(col: u16, footer_rows: &[u32], single_rows: &[u32]) -> String {
    let col_letter = col_to_letter(col);
    let refs: Vec<String> = footer_rows
        .iter()
        .chain(single_rows.iter())
        .map(|r| format!("{}{}", col_letter, r + 1))
        .collect();
    format!("=SUM({})", refs.join("+"))
}

/// Konvertiert Spalten-Index zu Buchstabe (0=A, 1=B, ...)
fn col_to_letter(col: u16) -> char {
    (b'A' + col as u8) as char
}

// ============================================================================
// Schreib-Helfer
// ============================================================================

/// Schreibt String mit Format
fn write_with_format(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    value: &str,
) -> Result<(), XlsxError> {
    if let Some(format) = fmt.get(row, col) {
        ws.write_string_with_format(row, col, value, format)?;
    } else {
        ws.write_string(row, col, value)?;
    }
    Ok(())
}

/// Schreibt Formel mit Format (locked)
fn write_formula_with_format(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    formula: &str,
) -> Result<(), XlsxError> {
    let formula = Formula::new(formula);
    if let Some(format) = fmt.get_locked(row, col) {
        ws.write_formula_with_format(row, col, formula, &format)?;
    } else {
        ws.write_formula(row, col, formula)?;
    }
    Ok(())
}

/// Schreibt Blank mit Format
fn write_blank_with_format(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
) -> Result<(), XlsxError> {
    if let Some(format) = fmt.get(row, col) {
        ws.write_blank(row, col, format)?;
    }
    Ok(())
}

/// Schreibt API-Werte für eine Position (C, D, E, F, H - nicht G!)
fn write_position_values(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
    values: &ReportValues,
    category: u8,
    position: u16,
) -> Result<(), XlsxError> {
    for field in PositionField::all() {
        let key = ApiKey::Position {
            category,
            position,
            field,
        };
        let value = values.get(key);

        if let Some(addr) = layout.position_addr(category, position, field) {
            write_cell_value(ws, fmt, addr.row, addr.col, value)?;
        }
    }
    Ok(())
}

/// Schreibt API-Werte für eine Single-Row Kategorie (D, E, F, H - nicht C, G!)
fn write_single_row_values(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
    values: &ReportValues,
    category: u8,
) -> Result<(), XlsxError> {
    for field in SingleRowField::all() {
        let key = ApiKey::SingleRow { category, field };
        let value = values.get(key);

        if let Some(addr) = layout.single_row_addr(category, field) {
            write_cell_value(ws, fmt, addr.row, addr.col, value)?;
        }
    }
    Ok(())
}

/// Schreibt einen CellValue mit Format
fn write_cell_value(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    value: &CellValue,
) -> Result<(), XlsxError> {
    match value {
        CellValue::Empty => {
            // Blank mit Format
            if let Some(format) = fmt.get(row, col) {
                ws.write_blank(row, col, format)?;
            }
        }
        CellValue::Text(s) => {
            if let Some(format) = fmt.get(row, col) {
                ws.write_string_with_format(row, col, s, format)?;
            } else {
                ws.write_string(row, col, s)?;
            }
        }
        CellValue::Number(n) => {
            if let Some(format) = fmt.get(row, col) {
                ws.write_number_with_format(row, col, *n, format)?;
            } else {
                ws.write_number(row, col, *n)?;
            }
        }
        CellValue::Date(d) => {
            // Date als String schreiben (mit Datumsformat)
            if let Some(format) = fmt.get(row, col) {
                ws.write_string_with_format(row, col, d, format)?;
            } else {
                ws.write_string(row, col, d)?;
            }
        }
    }
    Ok(())
}

/// Schreibt Formel in gemergten Bereich (col_start:col_end)
fn write_merged_formula(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col_start: u16,
    col_end: u16,
    formula: &str,
) -> Result<(), XlsxError> {
    let formula = Formula::new(formula);
    if let Some(format) = fmt.get_locked(row, col_start) {
        ws.merge_range(row, col_start, row, col_end, "", &format)?;
        ws.write_formula_with_format(row, col_start, formula, &format)?;
    } else {
        ws.merge_range(row, col_start, row, col_end, "", &Format::new())?;
        ws.write_formula(row, col_start, formula)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_vlookup_formula() {
        let formula = vlookup_formula(29);
        assert!(formula.contains("29"));
        assert!(formula.contains("VLOOKUP"));
    }

    #[test]
    fn test_sumproduct_formula() {
        let formula = sumproduct_formula(3, 27, 46); // D28:D47 (0-basiert: 27-46)
        assert_eq!(formula, "=SUMPRODUCT(ROUND(D28:D47,2))");
    }

    #[test]
    fn test_sum_formula() {
        let formula = sum_formula(3, &[47, 69], &[91, 92]);
        assert_eq!(formula, "=SUM(D48+D70+D92+D93)");
    }

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), 'A');
        assert_eq!(col_to_letter(1), 'B');
        assert_eq!(col_to_letter(3), 'D');
        assert_eq!(col_to_letter(7), 'H');
    }
}
