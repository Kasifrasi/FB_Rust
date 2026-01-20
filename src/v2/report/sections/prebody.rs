//! Pre-Body Section (Zeilen 20-25)
//!
//! Statischer Bereich zwischen Einnahme-Tabelle und dynamischem Body.
//! Enthält:
//! - Spacer Row 20-21
//! - Footer-Tabelle (Rows 22-25) mit vertikalen Merges

use crate::v2::report::ReportStyles;
use rust_xlsxwriter::{FormatAlign, Formula, Worksheet, XlsxError};

/// Schreibt die Pre-Body Section (Zeilen 20-25)
pub fn write_prebody_section(ws: &mut Worksheet, styles: &ReportStyles) -> Result<(), XlsxError> {
    // Spacer Row 20 (0-basiert: 20 = Excel 21)
    ws.set_row_height(20, 13.5)?;

    // Footer-Tabelle (Rows 22-25, 0-basiert: 22-25)
    write_footer_table(ws, styles)?;

    Ok(())
}

/// Schreibt die Footer-Tabelle (Zeilen 22-25)
fn write_footer_table(ws: &mut Worksheet, styles: &ReportStyles) -> Result<(), XlsxError> {
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

    // D23:D26 merged - Referenz auf D11
    ws.merge_range(22, 3, 25, 3, "", &fmt_ft_val)?;
    ws.write_formula_with_format(22, 3, Formula::new("=D11"), &fmt_ft_val)?;

    // E23:E26 merged - VLOOKUP(25) "Ausgaben"
    ws.merge_range(22, 4, 25, 4, "", &fmt_ft_val_bold)?;
    ws.write_formula_with_format(
        22,
        4,
        Formula::new(r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,25,FALSE))"#),
        &fmt_ft_val_bold,
    )?;

    // F23:F26 merged - VLOOKUP(55)
    ws.merge_range(22, 5, 25, 5, "", &fmt_ft_val)?;
    ws.write_formula_with_format(
        22,
        5,
        Formula::new(r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,55,FALSE))"#),
        &fmt_ft_val,
    )?;

    // G23:G26 merged - VLOOKUP(56)
    ws.merge_range(22, 6, 25, 6, "", &fmt_ft_val)?;
    ws.write_formula_with_format(
        22,
        6,
        Formula::new(r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,56,FALSE))"#),
        &fmt_ft_val,
    )?;

    // H23:H26 merged - Referenz auf H11
    ws.merge_range(22, 7, 25, 7, "", &fmt_ft_right)?;
    ws.write_formula_with_format(22, 7, Formula::new("=H11"), &fmt_ft_right)?;

    // B23, C23 - Blanks
    ws.write_blank(22, 1, &fmt_ft_lbl_b)?;
    ws.write_blank(22, 2, &fmt_ft_lbl_c)?;

    // === Row 23 (Excel 24): B24:C24 merged - VLOOKUP(24) ===
    ws.merge_range(23, 1, 23, 2, "", &fmt_ft_mid_bold)?;
    ws.write_formula_with_format(
        23,
        1,
        Formula::new(r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,24,FALSE))"#),
        &fmt_ft_mid_bold,
    )?;

    // === Row 24 (Excel 25): B25:C25 merged - Referenz auf B13 ===
    ws.merge_range(24, 1, 24, 2, "", &fmt_ft_mid)?;
    ws.write_formula_with_format(24, 1, Formula::new("=B13"), &fmt_ft_mid)?;

    // === Row 25 (Excel 26): B26:C26 merged mit thin bottom border ===
    ws.merge_range(25, 1, 25, 2, "", &fmt_ft_bot_bc)?;

    Ok(())
}
