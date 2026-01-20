//! Layout-Konfiguration für den Finanzbericht
//!
//! Enthält:
//! - Spaltenbreiten
//! - Zeilenhöhen
//! - Merge-Bereiche
//! - Freeze Panes
//! - Page Setup

use rust_xlsxwriter::{Format, Worksheet, XlsxError};

// ============================================================================
// Merge-Definition
// ============================================================================

pub struct MergeRange {
    pub first_row: u32,
    pub first_col: u16,
    pub last_row: u32,
    pub last_col: u16,
}

impl MergeRange {
    pub const fn new(first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> Self {
        Self {
            first_row,
            first_col,
            last_row,
            last_col,
        }
    }
}

// ============================================================================
// Sheet Setup (Spaltenbreiten, Zeilenhöhen, Zoom, etc.)
// ============================================================================

pub fn setup_sheet(ws: &mut Worksheet) -> Result<(), XlsxError> {
    // Page Setup
    ws.set_active(true);
    ws.set_zoom(85);
    ws.set_screen_gridlines(false);

    // Spaltenbreiten
    setup_column_widths(ws)?;

    // Zeilenhöhen
    setup_row_heights(ws)?;

    Ok(())
}

fn setup_column_widths(ws: &mut Worksheet) -> Result<(), XlsxError> {
    ws.set_column_width(0, 2.60)?; // A
    ws.set_column_width(1, 4.10)?; // B
    ws.set_column_width(2, 56.0)?; // C
    ws.set_column_width(3, 20.0)?; // D
    ws.set_column_width(4, 18.0)?; // E
    ws.set_column_width(5, 18.0)?; // F
    ws.set_column_width(6, 12.0)?; // G
    ws.set_column_width(7, 18.0)?; // H
    ws.set_column_width(8, 4.0)?; // I
    ws.set_column_width(9, 4.0)?; // J
    ws.set_column_width(10, 19.14)?; // K
    ws.set_column_width(11, 10.85)?; // L
    ws.set_column_width(12, 15.0)?; // M
    ws.set_column_width(13, 15.0)?; // N
    ws.set_column_width(14, 15.0)?; // O
    ws.set_column_width(15, 5.0)?; // P
    ws.set_column_width(16, 4.0)?; // Q
    ws.set_column_width(17, 19.14)?; // R
    ws.set_column_width(18, 10.85)?; // S
    ws.set_column_width(19, 15.0)?; // T
    ws.set_column_width(20, 15.0)?; // U
    ws.set_column_width(21, 15.0)?; // V
    ws.set_column_width(22, 4.0)?; // W
    ws.set_column_width(23, 36.71)?; // X
    ws.set_column_width(24, 10.85)?; // Y
    Ok(())
}

fn setup_row_heights(ws: &mut Worksheet) -> Result<(), XlsxError> {
    ws.set_row_height(1, 15.0)?;
    ws.set_row_height(2, 15.0)?;
    ws.set_row_height(3, 12.0)?;
    ws.set_row_height(4, 15.0)?;
    ws.set_row_height(5, 15.0)?;
    ws.set_row_height(6, 15.0)?;
    ws.set_row_height(7, 15.0)?;
    ws.set_row_height(8, 15.0)?;
    ws.set_row_height(9, 13.5)?;
    ws.set_row_height(10, 12.6)?;
    Ok(())
}

// ============================================================================
// Freeze Panes
// ============================================================================

pub fn setup_freeze_panes(ws: &mut Worksheet, row: u32) -> Result<(), XlsxError> {
    ws.set_freeze_panes(row, 0)?;
    Ok(())
}

// ============================================================================
// Default Column Format (unlocked)
// ============================================================================

pub fn setup_default_column_format(
    ws: &mut Worksheet,
    unlocked_format: &Format,
) -> Result<(), XlsxError> {
    for col in 0..1000u16 {
        ws.set_column_format(col, unlocked_format)?;
    }
    Ok(())
}

// ============================================================================
// Column Hiding
// ============================================================================

/// Hides columns Q through V (cols 16-21, 0-indexed)
///
/// The right panel has duplicate data in Q:V that is typically hidden
/// from end users while still being available for formulas.
pub fn hide_columns_qv(ws: &mut Worksheet) -> Result<(), XlsxError> {
    for col in 16..=21u16 {
        ws.set_column_hidden(col)?;
    }
    Ok(())
}

// ============================================================================
// Merge-Bereiche für Header
// ============================================================================

pub const HEADER_MERGES: &[MergeRange] = &[
    MergeRange::new(0, 1, 0, 2),  // B1:C1
    MergeRange::new(1, 1, 1, 2),  // B2:C2
    MergeRange::new(1, 4, 1, 7),  // E2:H2
    MergeRange::new(1, 9, 2, 14), // J2:O3
    MergeRange::new(2, 1, 2, 2),  // B3:C3
    MergeRange::new(3, 9, 3, 14), // J4:O4
    MergeRange::new(4, 1, 4, 2),  // B5:C5
    MergeRange::new(5, 1, 6, 2),  // B6:C7
    MergeRange::new(5, 3, 6, 7),  // D6:H7
    MergeRange::new(7, 1, 7, 2),  // B8:C8
    MergeRange::new(7, 6, 7, 7),  // G8:H8
    MergeRange::new(8, 1, 8, 2),  // B9:C9
    MergeRange::new(8, 6, 8, 7),  // G9:H9
];

pub const TABLE_HEADER_MERGES: &[MergeRange] = &[
    MergeRange::new(10, 1, 10, 2), // B11:C11
    MergeRange::new(10, 3, 13, 3), // D11:D14
    MergeRange::new(10, 4, 13, 4), // E11:E14
    MergeRange::new(10, 5, 13, 5), // F11:F14
    MergeRange::new(10, 6, 13, 6), // G11:G14
    MergeRange::new(10, 7, 13, 7), // H11:H14
    MergeRange::new(11, 1, 11, 2), // B12:C12
    MergeRange::new(12, 1, 12, 2), // B13:C13
    MergeRange::new(13, 1, 13, 2), // B14:C14
];

pub const TABLE_BODY_MERGES: &[MergeRange] = &[
    MergeRange::new(14, 1, 14, 2), // B15:C15
    MergeRange::new(15, 1, 15, 2), // B16:C16
    MergeRange::new(16, 1, 16, 2), // B17:C17
    MergeRange::new(17, 1, 17, 2), // B18:C18
    MergeRange::new(18, 1, 18, 2), // B19:C19
    MergeRange::new(19, 1, 19, 2), // B20:C20
];

pub const RIGHT_PANEL_MERGES: &[MergeRange] = &[
    MergeRange::new(10, 9, 10, 10),  // J11:K11
    MergeRange::new(10, 16, 10, 17), // Q11:R11
];
