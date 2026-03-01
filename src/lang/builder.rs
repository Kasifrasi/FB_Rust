use super::data::TEXT_MATRIX;
use crate::report::api::Currency;
use rust_xlsxwriter::{Format, Workbook, XlsxError};

/// Erstellt das Sprachversionen-Sheet (sichtbar)
pub fn build_sheet(workbook: &mut Workbook) -> Result<(), XlsxError> {
    build_sheet_with_visibility(workbook, false)
}

/// Erstellt das Sprachversionen-Sheet mit optionaler Ausblendung
///
/// # Arguments
/// * `workbook` - Das Workbook
/// * `hidden` - Wenn true, wird das Sheet ausgeblendet
pub fn build_sheet_with_visibility(workbook: &mut Workbook, hidden: bool) -> Result<(), XlsxError> {
    let ws = workbook.add_worksheet().set_name("Sprachversionen")?;

    // Sheet ausblenden wenn gewünscht
    if hidden {
        ws.set_hidden(true);
    }

    let arial = Format::new().set_font_name("Arial").set_font_size(10.0);

    // Column A (0): Currencies
    for (row_idx, currency) in Currency::all().iter().enumerate() {
        ws.write_string_with_format(row_idx as u32, 0, (*currency).as_str(), &arial)?;
    }

    // Columns B onwards: TEXT_MATRIX[lang_idx][term_idx] → Sheet[lang_idx][term_idx + 1]
    for (row_idx, row) in TEXT_MATRIX.iter().enumerate() {
        for (col_idx, value) in row.iter().enumerate() {
            ws.write_string_with_format(row_idx as u32, (col_idx + 1) as u16, *value, &arial)?;
        }
    }

    Ok(())
}
