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
//!
//! ## Formel-Evaluierung
//!
//! Alle Formeln werden mit gecachten Ergebnissen geschrieben:
//! - Ratio (G): F / D mit IFERROR(0)
//! - SUMPRODUCT: Summe der gerundeten Werte
//! - SUM (Total): Summe aller Kategorie-Summen

use std::collections::HashMap;

use super::config::BodyConfig;
use super::layout::{BodyLayout, CategoryLayout, CategoryMode, PositionRange, TOTAL_LABEL_INDEX};
use crate::report::api::{ApiKey, CellValue, PositionField, ReportValues};
use crate::report::core::lookup_text_string;
use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Format, Formula, Worksheet, XlsxError};

/// Gecachte Formel-Ergebnisse für den Body-Bereich
#[derive(Debug, Default)]
pub struct BodyFormulaCache {
    /// Ratio-Werte (G-Spalte): row -> value
    pub ratios: HashMap<u32, f64>,
    /// SUMPRODUCT-Werte (Footer D, E, F): (row, col) -> value
    pub sumproducts: HashMap<(u32, u16), f64>,
    /// Total SUM-Werte (Total D, E, F): col -> value
    pub totals: HashMap<u16, f64>,
    /// VLOOKUP-Texte: (row, col) -> text
    pub vlookups: HashMap<(u32, u16), String>,
}

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

/// Ratio-Formel für G-Spalte
const RATIO_FORMULA: &str = "=IFERROR(INDEX($F$1:$F$1001,ROW())/INDEX($D$1:$D$1001,ROW()),0)";

// ============================================================================
// Formel-Evaluierung
// ============================================================================

/// Evaluiert alle Body-Formeln und gibt gecachte Ergebnisse zurück
///
/// Evaluiert in der richtigen Reihenfolge:
/// 1. VLOOKUP-Texte für Labels (Kategorien, Footer, Total)
/// 2. Ratio-Formeln für Positions-Zeilen und Header-Input (G = F / D)
/// 3. SUMPRODUCT-Formeln für Footer (D, E, F = Summe der Positionen)
/// 4. Ratio-Formeln für Footer (G = F / D des Footers)
/// 5. SUM-Formeln für Total (D, E, F = Summe aller Kategorien)
/// 6. Ratio-Formel für Total (G = F / D des Totals)
pub fn evaluate_body_formulas(layout: &BodyLayout, values: &ReportValues) -> BodyFormulaCache {
    let mut cache = BodyFormulaCache::default();
    let language = values.language();

    // Sammle alle Werte für D, E, F pro Zeile
    let mut row_values: HashMap<u32, (f64, f64, f64)> = HashMap::new();

    for cat in &layout.categories {
        match &cat.mode {
            CategoryMode::HeaderInput { row } => {
                // VLOOKUP für Header-Input Label (C-Spalte)
                if let Some(text) = lookup_text_string(language, cat.meta.label_index) {
                    cache.vlookups.insert((*row, 2), text);
                }

                // Header-Input: Werte direkt aus ReportValues
                let d = get_position_value(values, cat.meta.num, 0, PositionField::Approved);
                let e = get_position_value(values, cat.meta.num, 0, PositionField::IncomeReport);
                let f = get_position_value(values, cat.meta.num, 0, PositionField::IncomeTotal);
                row_values.insert(*row, (d, e, f));

                // Ratio für Header-Input
                let ratio = calculate_ratio(f, d);
                cache.ratios.insert(*row, ratio);
            }
            CategoryMode::WithPositions {
                header_row,
                positions,
                footer_row,
            } => {
                // VLOOKUP für Header Label (C-Spalte)
                if let Some(text) = lookup_text_string(language, cat.meta.label_index) {
                    cache.vlookups.insert((*header_row, 2), text);
                }

                // VLOOKUP für Footer Sum-Label (B:C merged, col 1)
                if let Some(text) = lookup_text_string(language, cat.meta.sum_label_index) {
                    cache.vlookups.insert((*footer_row, 1), text);
                }

                // Positions-Zeilen: Werte aus ReportValues
                let mut sum_d = 0.0;
                let mut sum_e = 0.0;
                let mut sum_f = 0.0;

                for i in 0..positions.count {
                    let row = positions.start_row + i as u32;
                    let pos = i + 1; // 1-basiert

                    let d = get_position_value(values, cat.meta.num, pos, PositionField::Approved);
                    let e =
                        get_position_value(values, cat.meta.num, pos, PositionField::IncomeReport);
                    let f =
                        get_position_value(values, cat.meta.num, pos, PositionField::IncomeTotal);

                    row_values.insert(row, (d, e, f));

                    // Ratio für Position
                    let ratio = calculate_ratio(f, d);
                    cache.ratios.insert(row, ratio);

                    // Summiere für SUMPRODUCT (mit Rundung auf 2 Dezimalstellen)
                    sum_d += round2(d);
                    sum_e += round2(e);
                    sum_f += round2(f);
                }

                // SUMPRODUCT für Footer
                cache.sumproducts.insert((*footer_row, 3), sum_d); // D
                cache.sumproducts.insert((*footer_row, 4), sum_e); // E
                cache.sumproducts.insert((*footer_row, 5), sum_f); // F

                // Speichere Footer-Werte für Total-Berechnung
                row_values.insert(*footer_row, (sum_d, sum_e, sum_f));

                // Ratio für Footer
                let ratio = calculate_ratio(sum_f, sum_d);
                cache.ratios.insert(*footer_row, ratio);
            }
        }
    }

    // VLOOKUP für Total Label (B:C merged, col 1)
    if let Some(text) = lookup_text_string(language, TOTAL_LABEL_INDEX) {
        cache.vlookups.insert((layout.total_row, 1), text);
    }

    // Total: Summe aller Kategorie-Summen
    let mut total_d = 0.0;
    let mut total_e = 0.0;
    let mut total_f = 0.0;

    for cat in &layout.categories {
        let sum_row = cat.sum_row();
        if let Some((d, e, f)) = row_values.get(&sum_row) {
            total_d += d;
            total_e += e;
            total_f += f;
        }
    }

    cache.totals.insert(3, total_d); // D
    cache.totals.insert(4, total_e); // E
    cache.totals.insert(5, total_f); // F

    // Ratio für Total
    let ratio = calculate_ratio(total_f, total_d);
    cache.ratios.insert(layout.total_row, ratio);

    cache
}

/// Holt den Wert eines Position-Feldes aus ReportValues
fn get_position_value(
    values: &ReportValues,
    category: u8,
    position: u16,
    field: PositionField,
) -> f64 {
    let key = ApiKey::Position {
        category,
        position,
        field,
    };
    match values.get(key) {
        CellValue::Number(n) => *n,
        _ => 0.0,
    }
}

/// Berechnet Ratio: F / D mit IFERROR(0)
fn calculate_ratio(f: f64, d: f64) -> f64 {
    if d == 0.0 {
        0.0
    } else {
        f / d
    }
}

/// Rundet auf 2 Dezimalstellen (für SUMPRODUCT ROUND)
fn round2(value: f64) -> f64 {
    (value * 100.0).round() / 100.0
}

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

    // 2. Formeln evaluieren (wenn Werte vorhanden)
    let cache = values
        .map(|v| evaluate_body_formulas(&layout, v))
        .unwrap_or_default();

    // 3. Kategorien schreiben
    for cat in &layout.categories {
        write_category(ws, fmt, cat, &layout, values, &cache)?;
    }

    // 4. Gesamt-Zeile schreiben
    write_total_row(ws, fmt, &layout, &cache)?;

    // 5. Ratio-Formeln anwenden (mit gecachten Ergebnissen)
    write_ratio_formulas(ws, fmt, &layout, &cache)?;

    // E-Total und F-Total aus Cache holen (col 4 = E, col 5 = F)
    let e_total = cache.totals.get(&4).copied();
    let f_total = cache.totals.get(&5).copied();

    Ok(BodyResult {
        layout: layout.clone(),
        last_row: layout.last_row,
        total_row: layout.total_row,
        e_total,
        f_total,
    })
}

/// Schreibt eine Kategorie (einheitlich für beide Modi)
fn write_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
    layout: &BodyLayout,
    values: Option<&ReportValues>,
    cache: &BodyFormulaCache,
) -> Result<(), XlsxError> {
    match &cat.mode {
        CategoryMode::HeaderInput { row } => {
            write_header_input_category(ws, fmt, cat, *row, layout, values, cache)?;
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
                cache,
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
    cache: &BodyFormulaCache,
) -> Result<(), XlsxError> {
    // B: Kategorie-Nummer
    write_with_format(ws, fmt, row, 1, &format!("{}.", cat.meta.num))?;

    // C: VLOOKUP Label mit gecachtem Ergebnis
    let label_formula = vlookup_formula(cat.meta.label_index);
    let cached_text = cache.vlookups.get(&(row, 2)).cloned();
    write_formula_with_cached_text(ws, fmt, row, 2, &label_formula, cached_text)?;

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
    positions: &PositionRange,
    footer_row: u32,
    layout: &BodyLayout,
    values: Option<&ReportValues>,
    cache: &BodyFormulaCache,
) -> Result<(), XlsxError> {
    // === Header-Zeile ===
    // B: Kategorie-Nummer
    write_with_format(ws, fmt, header_row, 1, &format!("{}.", cat.meta.num))?;

    // C: VLOOKUP Label mit gecachtem Ergebnis
    let label_formula = vlookup_formula(cat.meta.label_index);
    let cached_text = cache.vlookups.get(&(header_row, 2)).cloned();
    write_formula_with_cached_text(ws, fmt, header_row, 2, &label_formula, cached_text)?;

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
    // B:C merged: VLOOKUP Sum-Label mit gecachtem Ergebnis
    let sum_label_formula = vlookup_formula(cat.meta.sum_label_index);
    let cached_sum_text = cache.vlookups.get(&(footer_row, 1)).cloned();
    write_merged_formula_with_cached_text(
        ws,
        fmt,
        footer_row,
        1,
        2,
        &sum_label_formula,
        cached_sum_text,
    )?;

    // D-F: SUMPRODUCT Formeln mit gecachten Ergebnissen
    for col in 3..=5 {
        let formula = sumproduct_formula(col, positions.start_row, positions.end_row);
        let cached_value = cache.sumproducts.get(&(footer_row, col)).copied();
        write_formula_with_cached_result(ws, fmt, footer_row, col, &formula, cached_value)?;
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
    cache: &BodyFormulaCache,
) -> Result<(), XlsxError> {
    let row = layout.total_row;

    // B:C merged: VLOOKUP "Gesamt" Label mit gecachtem Ergebnis
    let label_formula = vlookup_formula(TOTAL_LABEL_INDEX);
    let cached_text = cache.vlookups.get(&(row, 1)).cloned();
    write_merged_formula_with_cached_text(ws, fmt, row, 1, 2, &label_formula, cached_text)?;

    // D-F: SUM Formeln (summiert alle Summen-Zeilen) mit gecachten Ergebnissen
    // Summen-Zeile = Footer bei Positions-Modus, Header-Input bei Header-Eingabe
    let sum_rows: Vec<u32> = layout.categories.iter().map(|c| c.sum_row()).collect();

    for col in 3..=5 {
        let formula = sum_formula(col, &sum_rows);
        let cached_value = cache.totals.get(&col).copied();
        write_formula_with_cached_result(ws, fmt, row, col, &formula, cached_value)?;
    }

    // G: Ratio (wird später geschrieben)
    // H: Blank
    write_blank_with_format(ws, fmt, row, 7)?;

    Ok(())
}

/// Schreibt alle Ratio-Formeln (G-Spalte) mit gecachten Ergebnissen
fn write_ratio_formulas(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
    cache: &BodyFormulaCache,
) -> Result<(), XlsxError> {
    for row in layout.ratio_rows() {
        let cached_value = cache.ratios.get(&row).copied();
        write_formula_with_cached_result(ws, fmt, row, 6, RATIO_FORMULA, cached_value)?;
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

/// Schreibt Formel mit Format und gecachtem Ergebnis (locked)
///
/// Das gecachte Ergebnis wird in die Excel-Formel als `result` eingebettet,
/// sodass Excel den Wert sofort anzeigt ohne neu zu berechnen.
fn write_formula_with_cached_result(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    formula_str: &str,
    cached_value: Option<f64>,
) -> Result<(), XlsxError> {
    let formula = if let Some(value) = cached_value {
        // set_result erwartet einen String, daher f64 -> String
        Formula::new(formula_str).set_result(value.to_string())
    } else {
        Formula::new(formula_str)
    };

    if let Some(format) = fmt.get_locked(row, col) {
        ws.write_formula_with_format(row, col, formula, &format)?;
    } else {
        ws.write_formula(row, col, formula)?;
    }
    Ok(())
}

/// Schreibt Formel mit Format und gecachtem Text-Ergebnis (locked)
///
/// Für VLOOKUP-Formeln die Text zurückgeben.
fn write_formula_with_cached_text(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col: u16,
    formula_str: &str,
    cached_text: Option<String>,
) -> Result<(), XlsxError> {
    let formula = if let Some(text) = cached_text {
        Formula::new(formula_str).set_result(text)
    } else {
        Formula::new(formula_str)
    };

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

/// Schreibt Formel in gemergten Bereich mit gecachtem Text-Ergebnis
fn write_merged_formula_with_cached_text(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    row: u32,
    col_start: u16,
    col_end: u16,
    formula_str: &str,
    cached_text: Option<String>,
) -> Result<(), XlsxError> {
    let formula = if let Some(text) = cached_text {
        Formula::new(formula_str).set_result(text)
    } else {
        Formula::new(formula_str)
    };

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
