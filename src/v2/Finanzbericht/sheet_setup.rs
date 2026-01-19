use rust_xlsxwriter::{Worksheet, XlsxError};

pub fn sheet_setup(ws: &mut Worksheet) -> Result<(), XlsxError> {
    // --- Page Setup ---
    ws.set_active(true);
    ws.set_zoom(85);
    ws.set_screen_gridlines(false);

    // --- Columns ---
    // A=0
    ws.set_column_width(0, 2.60)?;
    ws.set_column_width(1, 4.10)?;
    ws.set_column_width(2, 56.0)?;
    ws.set_column_width(3, 20.0)?;
    ws.set_column_width(4, 18.0)?;
    ws.set_column_width(5, 18.0)?;
    ws.set_column_width(6, 12.0)?;
    ws.set_column_width(7, 18.0)?;
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

    // --- Row Heights (Header) ---
    ws.set_row_height(3, 12.75)?;
    ws.set_row_height(4, 12.0)?;
    ws.set_row_height(9, 13.5)?;
    ws.set_row_height(10, 12.6)?;

    Ok(())
}
