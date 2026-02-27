//! Gemeinsame Hilfs-Funktionen für Section-Writer

use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Worksheet, XlsxError};

/// Schreibt eine Blank-Zelle mit Format aus FormatMatrix.
/// Wenn kein Format für die Zelle existiert, wird nichts geschrieben.
pub fn write_blank(
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
