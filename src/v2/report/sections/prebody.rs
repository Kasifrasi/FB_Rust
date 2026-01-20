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
//!
//! Für Backward-Kompatibilität existiert weiterhin die Legacy-Version
//! `write_prebody_section`, die die Formeln selbst schreibt.

use crate::v2::report::definitions::lookup_text_string;
use crate::v2::report::ReportStyles;
use rust_xlsxwriter::{FormatAlign, Formula, Worksheet, XlsxError};

/// Schreibt die Pre-Body Section (Zeilen 20-25) - Legacy-Version
///
/// Diese Version schreibt die VLOOKUP-Formeln selbst mit gecachten Ergebnissen.
/// Für die einheitliche Registry-basierte Version siehe `write_prebody_section_unified`.
///
/// # Arguments
/// * `ws` - Das Worksheet
/// * `styles` - Report-Styles
/// * `language` - Die Sprache für VLOOKUP-Evaluierung (z.B. Some("deutsch"))
pub fn write_prebody_section(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    // Spacer Row 20 (0-basiert: 20 = Excel 21)
    ws.set_row_height(20, 13.5)?;

    // Footer-Tabelle (Rows 22-25, 0-basiert: 22-25)
    write_footer_table(ws, styles, language)?;

    Ok(())
}

/// Schreibt die Pre-Body Section mit Registry-evaluierten Werten
///
/// Diese Version nutzt die zentrale Registry für die VLOOKUP-Evaluierung,
/// schreibt aber die Formeln selbst mit den korrekten Formaten (Borders, Bold).
///
/// Die Formeln sind in der Registry registriert (`definitions.rs::register_prebody_formulas`),
/// aber diese Funktion schreibt sie mit den spezifischen Formatierungen.
///
/// # Arguments
/// * `ws` - Das Worksheet
/// * `styles` - Report-Styles
/// * `language` - Die Sprache für VLOOKUP-Evaluierung
pub fn write_prebody_section_unified(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    // Spacer Row 20 (0-basiert: 20 = Excel 21)
    ws.set_row_height(20, 13.5)?;

    // Footer-Tabelle mit Formeln und korrekten Formaten
    write_footer_table(ws, styles, language)?;

    Ok(())
}

/// Schreibt die Footer-Tabelle (Zeilen 22-25)
fn write_footer_table(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    let thin = styles.border_thin;
    let medium = styles.border_medium;

    // === Formate für Footer-Tabelle ===

    // D-H: Vertikale Merge-Zellen (Spaltenüberschriften)
    let fmt_ft_val = styles
        .base
        .clone()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_top(medium)
        .set_border_left(thin)
        .set_border_right(thin);

    let fmt_ft_val_bold = fmt_ft_val.clone().set_bold();
    let fmt_ft_right = fmt_ft_val.clone().set_border_right(medium);

    // B-C: Zeilen-Labels (B hat keine rechte border, da merged mit C)
    let fmt_ft_lbl_b = styles
        .base
        .clone()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_left(medium)
        .set_border_top(medium);
    // Keine rechte border auf B!

    let fmt_ft_lbl_c = styles
        .base
        .clone()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_top(medium)
        .set_border_right(thin);

    // B24:C24 und B25:C25 merged - Format für merge
    let fmt_ft_mid = styles
        .base
        .clone()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_border_left(medium)
        .set_border_right(thin);

    let fmt_ft_mid_bold = fmt_ft_mid.clone().set_bold();

    // B26:C26 - thin bottom (nicht medium!)
    let fmt_ft_bot_bc = styles
        .base
        .clone()
        .set_border_left(medium)
        .set_border_bottom(thin)
        .set_border_right(thin);

    // === Row 22 (Excel 23): Spaltenüberschriften mit vertikalen Merges ===

    // D23:D26 merged - VLOOKUP(11)
    ws.merge_range(22, 3, 25, 3, "", &fmt_ft_val)?;
    let formula_11 = make_vlookup_formula(11, language);
    ws.write_formula_with_format(22, 3, formula_11, &fmt_ft_val)?;

    // E23:E26 merged - VLOOKUP(25) "Ausgaben"
    ws.merge_range(22, 4, 25, 4, "", &fmt_ft_val_bold)?;
    let formula_25 = make_vlookup_formula(25, language);
    ws.write_formula_with_format(22, 4, formula_25, &fmt_ft_val_bold)?;

    // F23:F26 merged - VLOOKUP(55)
    ws.merge_range(22, 5, 25, 5, "", &fmt_ft_val)?;
    let formula_55 = make_vlookup_formula(55, language);
    ws.write_formula_with_format(22, 5, formula_55, &fmt_ft_val)?;

    // G23:G26 merged - VLOOKUP(56)
    ws.merge_range(22, 6, 25, 6, "", &fmt_ft_val)?;
    let formula_56 = make_vlookup_formula(56, language);
    ws.write_formula_with_format(22, 6, formula_56, &fmt_ft_val)?;

    // H23:H26 merged - VLOOKUP(15)
    ws.merge_range(22, 7, 25, 7, "", &fmt_ft_right)?;
    let formula_15 = make_vlookup_formula(15, language);
    ws.write_formula_with_format(22, 7, formula_15, &fmt_ft_right)?;

    // B23, C23 - Blanks
    ws.write_blank(22, 1, &fmt_ft_lbl_b)?;
    ws.write_blank(22, 2, &fmt_ft_lbl_c)?;

    // === Row 23 (Excel 24): B24:C24 merged - VLOOKUP(24) ===
    ws.merge_range(23, 1, 23, 2, "", &fmt_ft_mid_bold)?;
    let formula_24 = make_vlookup_formula(24, language);
    ws.write_formula_with_format(23, 1, formula_24, &fmt_ft_mid_bold)?;

    // === Row 24 (Excel 25): B25:C25 merged - VLOOKUP(10) (Währung) ===
    ws.merge_range(24, 1, 24, 2, "", &fmt_ft_mid)?;
    let formula_10 = make_vlookup_formula(10, language);
    ws.write_formula_with_format(24, 1, formula_10, &fmt_ft_mid)?;

    // === Row 25 (Excel 26): B26:C26 merged mit thin bottom border ===
    ws.merge_range(25, 1, 25, 2, "", &fmt_ft_bot_bc)?;

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
