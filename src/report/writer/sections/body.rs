//! Body Writer - Schreibt die Struktur des dynamischen Body-Bereichs
//!
//! Schreibt nur statische Strukturelemente (Nummern, Blanks, Merges).
//! Alle Formeln und API-Werte werden von `write_cells_from_registry()` geschrieben.
//!
//! ## Struktur pro Kategorie
//!
//! ### Header-Eingabe-Modus (0 Positionen)
//! - B: Nummer (z.B. "6.")
//! - C-H: via Registry (VLOOKUP, API, Ratio)
//!
//! ### Positions-Modus (1+ Positionen)
//! - Header: B Nummer, D-H Blanks
//! - Positions: B Nummer (z.B. "1.1")
//! - Footer: B:C Merge, H Blank
//!
//! ### Gesamt-Zeile
//! - B:C Merge, H Blank

use super::utils::write_blank;
use crate::report::body::{BodyLayout, CategoryLayout, CategoryMode};
use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Worksheet, XlsxError};

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

/// Schreibt die Body-Struktur (nur Nummern, Blanks, Merges)
///
/// Alle Formeln und API-Werte werden von `write_cells_from_registry()` geschrieben.
pub fn write_body_structure(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
) -> Result<(), XlsxError> {
    for cat in &layout.categories {
        write_category(ws, fmt, cat)?;
    }

    // Gesamt-Zeile: B:C Merge + H Blank
    write_total_row(ws, fmt, layout.total_row)?;

    Ok(())
}

/// Schreibt eine Kategorie (einheitlich für beide Modi)
fn write_category(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &CategoryLayout,
) -> Result<(), XlsxError> {
    match &cat.mode {
        CategoryMode::HeaderInput { row } => {
            // B: Kategorie-Nummer
            write_with_format(ws, fmt, *row, 1, &format!("{}.", cat.meta.num))?;
        }
        CategoryMode::WithPositions {
            header_row,
            positions,
            footer_row,
        } => {
            // === Header-Zeile ===
            // B: Kategorie-Nummer
            write_with_format(ws, fmt, *header_row, 1, &format!("{}.", cat.meta.num))?;

            // D-H: Blanks (Header hat keine Eingaben)
            for col in 3..=7 {
                write_blank(ws, fmt, *header_row, col)?;
            }

            // === Positions-Zeilen ===
            for i in 0..positions.count {
                let row = positions.start_row + i as u32;
                let pos_num = i + 1;

                // B: Positions-Nummer
                write_with_format(ws, fmt, row, 1, &format!("{}.{}", cat.meta.num, pos_num))?;
            }

            // === Footer-Zeile ===
            // B:C Merge (VLOOKUP Sum-Label wird von Registry geschrieben)
            if let Some(format) = fmt.get_locked(*footer_row, 1) {
                ws.merge_range(*footer_row, 1, *footer_row, 2, "", &format)?;
            }

            // H: Blank
            write_blank(ws, fmt, *footer_row, 7)?;
        }
    }
    Ok(())
}

/// Schreibt die Gesamt-Zeile (nur Merge + Blank)
fn write_total_row(ws: &mut Worksheet, fmt: &FormatMatrix, row: u32) -> Result<(), XlsxError> {
    // B:C Merge (VLOOKUP "Gesamt" Label wird von Registry geschrieben)
    if let Some(format) = fmt.get_locked(row, 1) {
        ws.merge_range(row, 1, row, 2, "", &format)?;
    }

    // H: Blank
    write_blank(ws, fmt, row, 7)?;

    Ok(())
}

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
