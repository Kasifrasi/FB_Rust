//! Table Section (Zeilen 10-19)
//!
//! Enthält:
//! - Table Header (Row 10-13)
//! - Table Body (Row 14-18)
//! - Summary Row (Row 19)
//!
//! **Hinweis:** Formeln werden von `write_cells_from_registry()` geschrieben.
//! Dieser Section-Writer schreibt nur Layout (Merges, Blanks, Body-Werte).

use crate::v2::report::formats::FormatMatrix;
use crate::v2::report::layout::MergeRange;
use rust_xlsxwriter::{Worksheet, XlsxError};

/// Merge-Bereiche für Table Header
pub const HEADER_MERGES: &[(MergeRange, u32, u16)] = &[
    (MergeRange::new(10, 1, 10, 2), 10, 1), // B11:C11
    (MergeRange::new(10, 3, 13, 3), 10, 3), // D11:D14
    (MergeRange::new(10, 4, 13, 4), 10, 4), // E11:E14
    (MergeRange::new(10, 5, 13, 5), 10, 5), // F11:F14
    (MergeRange::new(10, 6, 13, 6), 10, 6), // G11:G14
    (MergeRange::new(10, 7, 13, 7), 10, 7), // H11:H14
    (MergeRange::new(11, 1, 11, 2), 11, 1), // B12:C12
    (MergeRange::new(12, 1, 12, 2), 12, 1), // B13:C13
    (MergeRange::new(13, 1, 13, 2), 13, 1), // B14:C14
];

/// Merge-Bereiche für Table Body
pub const BODY_MERGES: &[(MergeRange, u32, u16)] = &[
    (MergeRange::new(14, 1, 14, 2), 14, 1), // B15:C15
    (MergeRange::new(15, 1, 15, 2), 15, 1), // B16:C16
    (MergeRange::new(16, 1, 16, 2), 16, 1), // B17:C17
    (MergeRange::new(17, 1, 17, 2), 17, 1), // B18:C18
    (MergeRange::new(18, 1, 18, 2), 18, 1), // B19:C19
    (MergeRange::new(19, 1, 19, 2), 19, 1), // B20:C20
];

/// Blank-Zellen für Table Header
pub const HEADER_BLANKS: &[(u32, u16)] = &[
    (10, 1),
    (10, 2), // B11, C11
    (13, 1),
    (13, 2), // B14, C14
];

/// Blank-Zellen für Table Body (Input/Calc/Pct/Right columns)
pub const BODY_BLANKS: &[(u32, u16)] = &[
    (14, 4),
    (14, 5),
    (14, 6),
    (14, 7),
    (15, 4),
    (15, 5),
    (15, 7),
    (16, 4),
    (16, 5),
    (16, 7),
    (17, 4),
    (17, 5),
    (17, 7),
    (18, 4),
    (18, 5),
    (18, 7),
    (19, 7), // Summary Row H20
];

/// Schreibt die Table Section (Layout, Merges, Blanks)
///
/// **Hinweis:** Formeln werden von `write_cells_from_registry()` geschrieben,
/// nicht hier. Die Registry enthält alle Formeln mit korrekten Evaluierungen.
pub fn write_table_section(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // Header Merges
    write_merges(ws, fmt, HEADER_MERGES)?;

    // Body Merges
    write_merges(ws, fmt, BODY_MERGES)?;

    // Header Blanks
    write_blanks(ws, fmt, HEADER_BLANKS)?;

    // Body Blanks
    write_blanks(ws, fmt, BODY_BLANKS)?;

    // Body Values (Zahlen)
    write_body_values(ws, fmt)?;

    // Formeln werden von write_cells_from_registry() geschrieben!
    // Die Registry enthält alle VLOOKUP- und Berechnungsformeln.

    Ok(())
}

fn write_merges(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    merges: &[(MergeRange, u32, u16)],
) -> Result<(), XlsxError> {
    for (range, fmt_row, fmt_col) in merges {
        if let Some(format) = fmt.get(*fmt_row, *fmt_col) {
            ws.merge_range(
                range.first_row,
                range.first_col,
                range.last_row,
                range.last_col,
                "",
                format,
            )?;
        }
    }
    Ok(())
}

fn write_blanks(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    blanks: &[(u32, u16)],
) -> Result<(), XlsxError> {
    for (row, col) in blanks {
        if let Some(format) = fmt.get(*row, *col) {
            ws.write_blank(*row, *col, format)?;
        }
    }
    Ok(())
}

fn write_body_values(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    // D15-D19: Default 0.0
    for row in 14..=18 {
        if let Some(format) = fmt.get(row, 3) {
            ws.write_number_with_format(row, 3, 0.0, format)?;
        }
    }
    Ok(())
}
