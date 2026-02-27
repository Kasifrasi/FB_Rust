//! Pre-Body Section (Zeilen 20-25)
//!
//! Statischer Bereich zwischen Einnahme-Tabelle und dynamischem Body.
//! Enthält:
//! - Spacer Row 20-21
//! - Footer-Tabelle (Rows 22-25) mit vertikalen Merges
//!
//! ## Registry-Integration
//!
//! Die VLOOKUP-Formeln in dieser Section sind jetzt in der zentralen
//! CellRegistry registriert (siehe `definitions.rs::register_prebody_formulas`).
//! Die Funktion `write_prebody_section` schreibt nur noch das Layout (Merges, Blanks),
//! während die Formeln von `write_cells_from_registry` geschrieben werden.

use crate::report::core::lookup_text_string;
use crate::report::format::FormatMatrix;
use rust_xlsxwriter::{Formula, Worksheet, XlsxError};

/// Schreibt die Pre-Body Section (Zeilen 20-25)
///
/// Formate werden aus der FormatMatrix gelesen — konsistent mit allen anderen
/// Section-Writern.
///
/// # Arguments
/// * `ws` - Das Worksheet
/// * `fmt` - FormatMatrix (befüllt durch `extend_format_matrix_with_prebody`)
/// * `language` - Die Sprache für VLOOKUP-Evaluierung (z.B. Some("deutsch"))
pub fn write_prebody_section(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    // Spacer Row 20 (0-basiert: 20 = Excel 21)
    ws.set_row_height(20, 13.5)?;

    // Footer-Tabelle (Rows 22-25, 0-basiert: 22-25)
    write_footer_table(ws, fmt, language)?;

    Ok(())
}

/// Schreibt die Footer-Tabelle (Zeilen 22-25)
fn write_footer_table(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    // === Row 22 (Excel 23): Spaltenüberschriften mit vertikalen Merges ===

    // D23:D26 merged - VLOOKUP(11)
    if let Some(format) = fmt.get_locked(22, 3) {
        ws.merge_range(22, 3, 25, 3, "", &format)?;
        ws.write_formula_with_format(22, 3, make_vlookup_formula(11, language), &format)?;
    }

    // E23:E26 merged - VLOOKUP(25) "Ausgaben" (bold)
    if let Some(format) = fmt.get_locked(22, 4) {
        ws.merge_range(22, 4, 25, 4, "", &format)?;
        ws.write_formula_with_format(22, 4, make_vlookup_formula(25, language), &format)?;
    }

    // F23:F26 merged - VLOOKUP(55)
    if let Some(format) = fmt.get_locked(22, 5) {
        ws.merge_range(22, 5, 25, 5, "", &format)?;
        ws.write_formula_with_format(22, 5, make_vlookup_formula(55, language), &format)?;
    }

    // G23:G26 merged - VLOOKUP(56)
    if let Some(format) = fmt.get_locked(22, 6) {
        ws.merge_range(22, 6, 25, 6, "", &format)?;
        ws.write_formula_with_format(22, 6, make_vlookup_formula(56, language), &format)?;
    }

    // H23:H26 merged - VLOOKUP(15)
    if let Some(format) = fmt.get_locked(22, 7) {
        ws.merge_range(22, 7, 25, 7, "", &format)?;
        ws.write_formula_with_format(22, 7, make_vlookup_formula(15, language), &format)?;
    }

    // B23, C23 - Blanks
    if let Some(format) = fmt.get(22, 1) {
        ws.write_blank(22, 1, format)?;
    }
    if let Some(format) = fmt.get(22, 2) {
        ws.write_blank(22, 2, format)?;
    }

    // === Row 23 (Excel 24): B24:C24 merged - VLOOKUP(24) ===
    if let Some(format) = fmt.get_locked(23, 1) {
        ws.merge_range(23, 1, 23, 2, "", &format)?;
        ws.write_formula_with_format(23, 1, make_vlookup_formula(24, language), &format)?;
    }

    // === Row 24 (Excel 25): B25:C25 merged - VLOOKUP(10) (Währung) ===
    if let Some(format) = fmt.get_locked(24, 1) {
        ws.merge_range(24, 1, 24, 2, "", &format)?;
        ws.write_formula_with_format(24, 1, make_vlookup_formula(10, language), &format)?;
    }

    // === Row 25 (Excel 26): B26:C26 merged blank (thin bottom border) ===
    if let Some(format) = fmt.get(25, 1) {
        ws.merge_range(25, 1, 25, 2, "", format)?;
    }

    Ok(())
}

/// Erstellt eine VLOOKUP-Formel mit gecachtem Text-Ergebnis
fn make_vlookup_formula(index: usize, language: Option<&str>) -> Formula {
    let formula_str = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
        index
    );

    if let Some(text) = lookup_text_string(language, index) {
        Formula::new(&formula_str).set_result(text)
    } else {
        Formula::new(&formula_str)
    }
}
