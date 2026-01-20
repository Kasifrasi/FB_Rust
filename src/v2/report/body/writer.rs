//! Body Writer - Schreibt die Struktur des dynamischen Body-Bereichs
//!
//! Schreibt für jede Kategorie basierend auf ihrem Modus:
//!
//! ## Header-Eingabe-Modus (0 Positionen)
//! - Eine Zeile: Nummer + VLOOKUP Label + API-Eingaben (D, E, F, H) + Ratio (G)
//!
//! ## Positions-Modus (1+ Positionen)
//! - Header-Zeile: Nummer + VLOOKUP Label
//! - N Positions-Zeilen: Nummer + API-Eingaben (C, D, E, F, H) + Ratio (G)
//! - Footer-Zeile: VLOOKUP Sum-Label + SUMPRODUCT Formeln + Ratio (G)
//!
//! Am Ende:
//! - Gesamt-Zeile: VLOOKUP Label + SUM Formeln + Ratio (G)

use super::config::BodyConfig;
use super::layout::{BodyLayout, CategoryLayout, CategoryMode, TOTAL_LABEL_INDEX};
use crate::v2::report::api::{ApiKey, PositionField};
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
        write_category(ws, fmt, cat, &layout, values)?;
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

/// Schreibt eine Kategorie (einheitlich für beide Modi)
fn write_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
    layout: &BodyLayout,
    values: Option<&ReportValues>,
) -> Result<(), XlsxError> {
    match &cat.mode {
        CategoryMode::HeaderInput { row } => {
            write_header_input_category(ws, fmt, cat, *row, layout, values)?;
        }
        CategoryMode::WithPositions {
            header_row,
            positions,
            footer_row,
        } => {
            write_positions_category(
                ws,
                fmt,
                cat,
                *header_row,
                positions,
                *footer_row,
                layout,
                values,
            )?;
        }
    }
    Ok(())
}

/// Schreibt eine Header-Eingabe-Kategorie (0 Positionen)
fn write_header_input_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
    row: u32,
    layout: &BodyLayout,
    values: Option<&ReportValues>,
) -> Result<(), XlsxError> {
    // B: Kategorie-Nummer
    write_with_format(ws, fmt, row, 1, &format!("{}.", cat.meta.num))?;

    // C: VLOOKUP Label
    let label_formula = vlookup_formula(cat.meta.label_index);
    write_formula_with_format(ws, fmt, row, 2, &label_formula)?;

    // D, E, F, H: API-Werte (position=0) oder Blanks
    // G wird später als Ratio-Formel geschrieben
    if let Some(values) = values {
        write_header_input_values(ws, fmt, layout, values, cat.meta.num)?;
    } else {
        // Fallback: Blanks für D, E, F, H
        for col in [3, 4, 5, 7] {
            write_blank_with_format(ws, fmt, row, col)?;
        }
    }

    Ok(())
}

/// Schreibt eine Positions-Kategorie (1+ Positionen)
fn write_positions_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
    header_row: u32,
    positions: &super::layout::PositionRange,
    footer_row: u32,
    layout: &BodyLayout,
    values: Option<&ReportValues>,
) -> Result<(), XlsxError> {
    // === Header-Zeile ===
    // B: Kategorie-Nummer
    write_with_format(ws, fmt, header_row, 1, &format!("{}.", cat.meta.num))?;

    // C: VLOOKUP Label
    let label_formula = vlookup_formula(cat.meta.label_index);
    write_formula_with_format(ws, fmt, header_row, 2, &label_formula)?;

    // D-H: Blanks (Header hat keine Eingaben)
    for col in 3..=7 {
        write_blank_with_format(ws, fmt, header_row, col)?;
    }

    // === Positions-Zeilen ===
    for i in 0..positions.count {
        let row = positions.start_row + i as u32;
        let pos_num = i + 1; // 1-basiert

        // B: Positions-Nummer
        write_with_format(ws, fmt, row, 1, &format!("{}.{}", cat.meta.num, pos_num))?;

        // C, D, E, F, H: API-Werte oder Blanks
        // G wird später als Ratio-Formel geschrieben
        if let Some(values) = values {
            write_position_values(ws, fmt, layout, values, cat.meta.num, pos_num)?;
        } else {
            // Fallback: Blanks
            for col in [2, 3, 4, 5, 7] {
                // C, D, E, F, H (nicht G)
                write_blank_with_format(ws, fmt, row, col)?;
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

    // D-F: SUM Formeln (summiert alle Summen-Zeilen)
    // Summen-Zeile = Footer bei Positions-Modus, Header-Input bei Header-Eingabe
    let sum_rows: Vec<u32> = layout.categories.iter().map(|c| c.sum_row()).collect();

    for col in 3..=5 {
        let formula = sum_formula(col, &sum_rows);
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
fn sum_formula(col: u16, sum_rows: &[u32]) -> String {
    let col_letter = col_to_letter(col);
    let refs: Vec<String> = sum_rows
        .iter()
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

/// Schreibt API-Werte für Header-Eingabe (position=0)
/// Felder: D, E, F, H (nicht C - ist VLOOKUP Label, nicht G - ist Ratio)
fn write_header_input_values(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
    values: &ReportValues,
    category: u8,
) -> Result<(), XlsxError> {
    // Bei Header-Eingabe: position=0, nur Felder ohne Description
    for field in PositionField::header_input_fields() {
        let key = ApiKey::Position {
            category,
            position: 0, // Header-Eingabe!
            field,
        };
        let value = values.get(key);

        if let Some(addr) = layout.position_addr(category, 0, field) {
            write_cell_value(ws, fmt, addr.row, addr.col, value)?;
        }
    }
    Ok(())
}

/// Schreibt API-Werte für eine Position (position >= 1)
/// Felder: C, D, E, F, H (nicht G - ist Ratio)
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
        let formula = sum_formula(3, &[47, 69, 91, 92]);
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
