//! Header Section (Zeilen 0-9)
//!
//! Enthält:
//! - Titel (B1)
//! - Sprache/Währung (E2, E3)
//! - Projekt-Infos (D5, D6)
//! - Projektlaufzeit/Berichtszeitraum (Row 7-8)
//! - Info-Box (J1-O4)

use crate::v2::lang::data::CURRENCIES;
use crate::v2::report::formats::FormatMatrix;
use crate::v2::report::layout::MergeRange;
use rust_xlsxwriter::{DataValidation, Worksheet, XlsxError};

/// Merge-Bereiche für Header Section
pub const MERGES: &[(MergeRange, u32, u16)] = &[
    (MergeRange::new(0, 1, 0, 2), 0, 1),  // B1:C1
    (MergeRange::new(1, 1, 1, 2), 1, 1),  // B2:C2
    (MergeRange::new(1, 9, 2, 14), 1, 9), // J2:O3
    (MergeRange::new(2, 1, 2, 2), 2, 1),  // B3:C3
    (MergeRange::new(3, 9, 3, 14), 3, 9), // J4:O4
    (MergeRange::new(4, 1, 4, 2), 4, 1),  // B5:C5
    (MergeRange::new(5, 1, 6, 2), 5, 1),  // B6:C7
    (MergeRange::new(5, 3, 6, 7), 5, 3),  // D6:H7
    (MergeRange::new(7, 1, 7, 2), 7, 1),  // B8:C8
    (MergeRange::new(7, 6, 7, 7), 7, 6),  // G8:H8
    (MergeRange::new(8, 1, 8, 2), 8, 1),  // B9:C9
    (MergeRange::new(8, 6, 8, 7), 8, 6),  // G9:H9
];

/// Blank-Zellen für Header Section
pub const BLANKS: &[(u32, u16)] = &[
    (0, 10),
    (0, 11),
    (0, 12),
    (0, 13),
    (0, 14), // Row 0
    (2, 4),  // E3
    (4, 3),  // D5
    (6, 9),  // J7
    (7, 4),
    (7, 9), // E8, J8
    (8, 4),
    (8, 6),
    (8, 9), // E9, G9, J9
];

/// Schreibt die Header Section
pub fn write_header_section(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    // Merges
    write_merges(ws, fmt)?;

    // Blanks
    write_blanks(ws, fmt)?;

    // Werte
    write_values(ws, fmt, suffix, lang_val)?;

    // Data Validations
    write_validations(ws)?;

    Ok(())
}

fn write_merges(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    for (range, fmt_row, fmt_col) in MERGES {
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

fn write_blanks(ws: &mut Worksheet, fmt: &FormatMatrix) -> Result<(), XlsxError> {
    for (row, col) in BLANKS {
        if let Some(format) = fmt.get(*row, *col) {
            ws.write_blank(*row, *col, format)?;
        }
    }
    Ok(())
}

fn write_values(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    suffix: &str,
    lang_val: &str,
) -> Result<(), XlsxError> {
    // Suffix (B2)
    if let Some(format) = fmt.get(1, 1) {
        ws.write_string_with_format(1, 1, suffix, format)?;
    }

    // Language (E2)
    if let Some(format) = fmt.get(1, 4) {
        ws.write_string_with_format(1, 4, lang_val, format)?;
    }

    Ok(())
}

fn write_validations(ws: &mut Worksheet) -> Result<(), XlsxError> {
    // E2: Sprache
    let lang_validation =
        DataValidation::new().allow_list_formula("=Sprachversionen!$B$1:$B$5".into());
    ws.add_data_validation(1, 4, 1, 4, &lang_validation)?;

    // E3: Währung
    let currency_count = CURRENCIES.len();
    let currency_formula = format!("=Sprachversionen!$A$1:$A${}", currency_count);
    let currency_validation =
        DataValidation::new().allow_list_formula(currency_formula.as_str().into());
    ws.add_data_validation(2, 4, 2, 4, &currency_validation)?;

    Ok(())
}
