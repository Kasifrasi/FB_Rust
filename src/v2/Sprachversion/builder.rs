use super::data::{CURRENCIES, TEXT_MATRIX};
use rust_xlsxwriter::{Workbook, XlsxError};

pub fn build_sheet(workbook: &mut Workbook) -> Result<(), XlsxError> {
    let ws = workbook.add_worksheet().set_name("Sprachversionen")?;

    // Column A (0): Currencies
    for (row_idx, currency) in CURRENCIES.iter().enumerate() {
        ws.write_string(row_idx as u32, 0, *currency)?;
    }

    // Columns B onwards (1+): Text Matrix (Languages)
    // Each row in TEXT_MATRIX corresponds to a language (Deutsch, English, etc.)
    // Each item in the inner vector corresponds to a translated term.
    // We map TEXT_MATRIX[lang_idx][term_idx] to Sheet[Row = lang_idx][Col = term_idx + 1]
    for (row_idx, row) in TEXT_MATRIX.iter().enumerate() {
        for (col_idx, value) in row.iter().enumerate() {
            ws.write_string(row_idx as u32, (col_idx + 1) as u16, *value)?;
        }
    }

    ws.set_hidden(true);
    Ok(())
}
