//! Layout-Konfiguration für den Finanzbericht
//!
//! Enthält:
//! - Spaltenbreiten
//! - Zeilenhöhen
//! - Merge-Bereiche
//! - Freeze Panes
//! - Page Setup

use rust_xlsxwriter::{Worksheet, XlsxError};

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
