//! Report Footer Sektion
//!
//! Der Footer kommt nach dem dynamischen Body-Bereich mit 3 Zeilen Abstand.
//! Basiert auf Template Zeilen 164-184.
//!
//! Struktur (21 Zeilen):
//! - Zeile 0: E: "Saldo für den Berichtszeitraum"
//! - Zeile 1: B: "ABSCHLUSS"
//! - Zeile 2: E: Währungs-Referenz (=B13)
//! - Zeile 3: Leer
//! - Zeile 4: B: "Saldo...", D: Check ✓, E: Differenz-Formel
//! - Zeile 5: Leer
//! - Zeile 6: B: "Saldenabstimmung:", E: OK-Check
//! - Zeile 7-9: B: "Bank/Kasse/Sonstiges", E: INPUT (API)
//! - Zeile 10-12: Leer
//! - Zeile 13: B: Bestätigung 1
//! - Zeile 14: B: Bestätigung 2
//! - Zeile 15-18: Leer
//! - Zeile 19: B: "Ort, Datum...", D: "Unterschrift..."
//! - Zeile 20: D: "Funktion..."

use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Formula, Worksheet, XlsxError};

use crate::v2::report::definitions::lookup_text_string;
use crate::v2::report::formats::ReportStyles;

/// Footer-Layout mit berechneten Zeilenpositionen
#[derive(Debug, Clone)]
pub struct FooterLayout {
    /// Startzeile des Footers (0-indexed)
    pub start_row: u32,
    /// Zeile für Saldo-Formel (E-Spalte)
    pub saldo_row: u32,
    /// Zeilen für Bank/Kasse/Sonstiges Input (E-Spalte)
    pub input_rows: [u32; 3],
    /// Letzte Zeile des Footers
    pub end_row: u32,
}

impl FooterLayout {
    /// Berechnet das Footer-Layout basierend auf der Total-Zeile des Body
    ///
    /// Der Footer beginnt 3 Zeilen nach dem Total.
    pub fn compute(total_row: u32) -> Self {
        let start_row = total_row + 3; // 3 Zeilen Abstand

        Self {
            start_row,
            saldo_row: start_row + 4, // Zeile 4: Saldo-Differenz
            input_rows: [
                start_row + 7, // Zeile 7: Bank
                start_row + 8, // Zeile 8: Kasse
                start_row + 9, // Zeile 9: Sonstiges
            ],
            end_row: start_row + 20, // 21 Zeilen (0-20)
        }
    }
}

/// Schreibt den Report-Footer
///
/// # Arguments
/// * `ws` - Das Worksheet
/// * `styles` - Report-Styles
/// * `total_row` - Die Total-Zeile des Body (0-indexed)
/// * `income_row` - Die Einnahmen-Zeile (Zeile 20, 0-indexed = 19)
/// * `language` - Die Sprache für VLOOKUP-Evaluierung (z.B. Some("deutsch"))
/// * `e_income` - E-Spalte Einnahmen (für Check-Formel)
/// * `e_total` - E-Spalte Total (für Check-Formel)
/// * `f_income` - F-Spalte Einnahmen (für Check-Formel)
/// * `f_total` - F-Spalte Total (für Check-Formel)
/// * `bank` - Bank-Wert (für OK-Formel)
/// * `kasse` - Kasse-Wert (für OK-Formel)
/// * `sonstiges` - Sonstiges-Wert (für OK-Formel)
///
/// # Returns
/// Das berechnete Footer-Layout
pub fn write_footer(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    total_row: u32,
    income_row: u32,
    language: Option<&str>,
    e_income: Option<f64>,
    e_total: Option<f64>,
    f_income: Option<f64>,
    f_total: Option<f64>,
    bank: Option<f64>,
    kasse: Option<f64>,
    sonstiges: Option<f64>,
) -> Result<FooterLayout, XlsxError> {
    let layout = FooterLayout::compute(total_row);
    let s = layout.start_row;

    // === BASIS-FORMATE ===
    // Formeln sind immer gesperrt (locked ist default)
    let base = Format::new().set_font_name("Arial").set_font_size(10.0);

    let normal = base
        .clone()
        .set_align(FormatAlign::Left)
        .set_align(FormatAlign::VerticalCenter);

    let bold = normal.clone().set_bold();

    let center = base
        .clone()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter);

    let center_bold = center.clone().set_bold();

    let center_wrap = center_bold.clone().set_text_wrap();

    // Grauer Text für OK-Check
    let gray_center = center
        .clone()
        .set_font_color(rust_xlsxwriter::Color::RGB(0x808080));

    // Zahlenformat für Saldo
    let number_right = base
        .clone()
        .set_align(FormatAlign::Right)
        .set_align(FormatAlign::VerticalCenter)
        .set_num_format("#,##0.00");

    // Input-Format mit Hintergrundfarbe (unlocked für Eingabe)
    let input_format = number_right
        .clone()
        .set_unlocked()
        .set_background_color(styles.fill_input);

    // === BORDER-FORMATE ===
    let border_medium = FormatBorder::Medium;
    let border_thin = FormatBorder::Thin;

    // =========================================================================
    // ZEILE 0 (s): "Kontrolle" / "Saldo für den Berichtszeitraum"
    // =========================================================================

    // B: leer mit top+left border
    let fmt_b_s0 = normal
        .clone()
        .set_border_top(border_medium)
        .set_border_left(border_medium);
    ws.write_blank(s, 1, &fmt_b_s0)?;

    // C: leer mit top border
    let fmt_c_s0 = normal.clone().set_border_top(border_medium);
    ws.write_blank(s, 2, &fmt_c_s0)?;

    // D: leer mit top border + right thin (Trennlinie zu E)
    let fmt_d_s0 = normal
        .clone()
        .set_border_top(border_medium)
        .set_border_right(border_thin);
    ws.write_blank(s, 3, &fmt_d_s0)?;

    // E(s):E(s+1) merged - Zuerst merge, dann Formel schreiben
    // Format für merged E: top+right border, zentriert, wrap
    let fmt_e_merged = center_wrap
        .clone()
        .set_border_top(border_medium)
        .set_border_right(border_medium)
        .set_border_left(border_thin);

    // E(s):E(s+1) merged mit VLOOKUP 44 ("Saldo für den Berichtszeitraum")
    write_merged_vlookup_formula(ws, s, 4, s + 1, 4, 44, &fmt_e_merged, language)?;

    // =========================================================================
    // ZEILE 1 (s+1): "ABSCHLUSS"
    // =========================================================================

    // B:D(s+1) merged mit VLOOKUP 43 ("ABSCHLUSS")
    let fmt_bcd_merged = bold
        .clone()
        .set_border_left(border_medium)
        .set_border_right(border_thin)
        .set_align(FormatAlign::Center);

    write_merged_vlookup_formula(ws, s + 1, 1, s + 1, 3, 43, &fmt_bcd_merged, language)?;
    // E ist bereits in merge von oben

    // =========================================================================
    // ZEILE 2 (s+2): Währungs-Referenz
    // =========================================================================

    // B: left border
    let fmt_b_s2 = normal.clone().set_border_left(border_medium);
    ws.write_blank(s + 2, 1, &fmt_b_s2)?;

    // C: zentriert (wie im Template)
    ws.write_blank(s + 2, 2, &center)?;

    // D: right thin border
    let fmt_d_s2 = normal.clone().set_border_right(border_thin);
    ws.write_blank(s + 2, 3, &fmt_d_s2)?;

    // E: VLOOKUP(10) (Währung), right border, left thin, zentriert
    let fmt_e_s2 = center
        .clone()
        .set_border_right(border_medium)
        .set_border_left(border_thin);
    write_vlookup_formula(ws, s + 2, 4, 10, &fmt_e_s2, language)?;

    // =========================================================================
    // ZEILE 3 (s+3): Leer
    // =========================================================================

    let fmt_b_s3 = normal.clone().set_border_left(border_medium);
    ws.write_blank(s + 3, 1, &fmt_b_s3)?;
    ws.write_blank(s + 3, 2, &normal)?;
    let fmt_d_s3 = normal.clone().set_border_right(border_thin);
    ws.write_blank(s + 3, 3, &fmt_d_s3)?;
    let fmt_e_s3 = normal
        .clone()
        .set_border_right(border_medium)
        .set_border_left(border_thin);
    ws.write_blank(s + 3, 4, &fmt_e_s3)?;

    // =========================================================================
    // ZEILE 4 (s+4): Saldo-Differenz
    // =========================================================================

    // B: VLOOKUP 45 ("Saldo..."), bold, left+top+bottom border
    let fmt_b_s4 = bold
        .clone()
        .set_border_left(border_medium)
        .set_border_top(border_thin)
        .set_border_bottom(border_thin);
    write_vlookup_formula(ws, s + 4, 1, 45, &fmt_b_s4, language)?;

    // C: top+bottom border
    let fmt_c_s4 = normal
        .clone()
        .set_border_top(border_thin)
        .set_border_bottom(border_thin);
    ws.write_blank(s + 4, 2, &fmt_c_s4)?;

    // D: Check-Formel ✓, zentriert, top+bottom border (KEIN right border - sonst erscheint er als E left!)
    let fmt_d_s4 = center
        .clone()
        .set_border_top(border_thin)
        .set_border_bottom(border_thin);
    let check_formula_str = format!(
        "=IF(ROUND(E{},2)=(ROUND(F{}-F{},2)),\"✓\",\"\")",
        s + 4 + 1,      // E saldo (1-indexed)
        income_row + 1, // F20 (1-indexed)
        total_row + 1   // F total (1-indexed)
    );
    // Evaluiere Check-Formel: Wenn ROUND(E_income - E_total, 2) == ROUND(F_income - F_total, 2) → "✓"
    // WICHTIG: Wenn Ergebnis leer ist, KEIN set_result() aufrufen - sonst zeigt Excel 0 statt leer!
    let check_formula = match (e_income, e_total, f_income, f_total) {
        (Some(e_inc), Some(e_tot), Some(f_inc), Some(f_tot)) => {
            let e_diff = (e_inc - e_tot) * 100.0;
            let f_diff = (f_inc - f_tot) * 100.0;
            if e_diff.round() == f_diff.round() {
                // Nur bei Erfolg set_result mit "✓"
                Formula::new(&check_formula_str).set_result("✓")
            } else {
                // Bei leerem Ergebnis: KEIN set_result - Excel evaluiert selbst
                Formula::new(&check_formula_str)
            }
        }
        // Keine Werte: kein set_result
        _ => Formula::new(&check_formula_str),
    };
    ws.write_formula_with_format(s + 4, 3, check_formula, &fmt_d_s4)?;

    // E: Differenz-Formel (E_income - E_total), right+top+bottom border (KEIN left border!)
    let fmt_e_s4 = number_right
        .clone()
        .set_border_right(border_medium)
        .set_border_top(border_thin)
        .set_border_bottom(border_thin);
    let diff_formula_str = format!("=E{}-E{}", income_row + 1, total_row + 1);
    // Evaluiere Differenz-Formel: E_income - E_total
    let diff_formula = match (e_income, e_total) {
        (Some(e_inc), Some(e_tot)) => {
            Formula::new(&diff_formula_str).set_result((e_inc - e_tot).to_string())
        }
        _ => Formula::new(&diff_formula_str),
    };
    ws.write_formula_with_format(s + 4, 4, diff_formula, &fmt_e_s4)?;

    // =========================================================================
    // ZEILE 5 (s+5): Leer
    // =========================================================================

    let fmt_b_s5 = normal.clone().set_border_left(border_medium);
    ws.write_blank(s + 5, 1, &fmt_b_s5)?;
    ws.write_blank(s + 5, 2, &normal)?;
    // D: KEIN right border (sonst erscheint er als E left!)
    ws.write_blank(s + 5, 3, &normal)?;
    // E: right border only (KEIN left border!)
    let fmt_e_s5 = normal.clone().set_border_right(border_medium);
    ws.write_blank(s + 5, 4, &fmt_e_s5)?;

    // =========================================================================
    // ZEILE 6 (s+6): Saldenabstimmung
    // =========================================================================

    // B: VLOOKUP 46 ("Saldenabstimmung:"), bold, left border
    let fmt_b_s6 = bold.clone().set_border_left(border_medium);
    write_vlookup_formula(ws, s + 6, 1, 46, &fmt_b_s6, language)?;

    // C: leer
    ws.write_blank(s + 6, 2, &normal)?;

    // D: KEIN right border (sonst erscheint er als E left!)
    ws.write_blank(s + 6, 3, &normal)?;

    // E: OK-Check, grau zentriert, right border only (KEIN left border!)
    let fmt_e_s6 = gray_center.clone().set_border_right(border_medium);
    let ok_formula_str = format!(
        "=IF(E{}=SUM(E{}:E{}),\"OK\",\"\")",
        s + 4 + 1, // saldo row (1-indexed)
        s + 7 + 1, // bank row (1-indexed)
        s + 9 + 1  // sonstiges row (1-indexed)
    );
    // Evaluiere OK-Formel: Wenn E_saldo == SUM(E_bank:E_sonstiges) → "OK"
    // E_saldo ist die Differenz E_income - E_total (wird in E s+4 geschrieben)
    // SUM(E_bank:E_sonstiges) = Bank + Kasse + Sonstiges
    // Wenn E_saldo == Bank + Kasse + Sonstiges → "OK"
    // WICHTIG: Wenn Ergebnis leer ist, KEIN set_result() aufrufen - sonst zeigt Excel 0 statt leer!
    let ok_formula = match (e_income, e_total, bank, kasse, sonstiges) {
        (Some(e_inc), Some(e_tot), Some(b), Some(k), Some(so)) => {
            let e_saldo = e_inc - e_tot;
            let sum_inputs = b + k + so;
            if (e_saldo * 100.0).round() == (sum_inputs * 100.0).round() {
                // Nur bei Erfolg set_result mit "OK"
                Formula::new(&ok_formula_str).set_result("OK")
            } else {
                // Bei leerem Ergebnis: KEIN set_result - Excel evaluiert selbst
                Formula::new(&ok_formula_str)
            }
        }
        // Keine Werte: kein set_result
        _ => Formula::new(&ok_formula_str),
    };
    ws.write_formula_with_format(s + 6, 4, ok_formula, &fmt_e_s6)?;

    // =========================================================================
    // ZEILEN 7-8 (s+7 bis s+8): Bank, Kasse
    // =========================================================================

    let vlookup_indices = [47, 48]; // Bank, Kasse
    for (i, vlookup_idx) in vlookup_indices.iter().enumerate() {
        let row = s + 7 + i as u32;

        // B: Label, left+top border
        let fmt_b = normal
            .clone()
            .set_border_left(border_medium)
            .set_border_top(border_thin);
        write_vlookup_formula(ws, row, 1, *vlookup_idx, &fmt_b, language)?;

        // C: top border
        let fmt_c = normal.clone().set_border_top(border_thin);
        ws.write_blank(row, 2, &fmt_c)?;

        // D: top+right thin border (KEIN left border!)
        let fmt_d = normal
            .clone()
            .set_border_top(border_thin)
            .set_border_right(border_thin);
        ws.write_blank(row, 3, &fmt_d)?;

        // E: Input-Feld, top+right+left border
        let fmt_e = input_format
            .clone()
            .set_border_top(border_thin)
            .set_border_right(border_medium)
            .set_border_left(border_thin);
        ws.write_blank(row, 4, &fmt_e)?;
    }

    // =========================================================================
    // ZEILE 9 (s+9): Sonstiges - letzte Zeile mit bottom border
    // =========================================================================

    // B: Label VLOOKUP 49 ("Sonstiges"), left+top+bottom border
    let fmt_b_s9 = normal
        .clone()
        .set_border_left(border_medium)
        .set_border_top(border_thin)
        .set_border_bottom(border_medium);
    write_vlookup_formula(ws, s + 9, 1, 49, &fmt_b_s9, language)?;

    // C: top+bottom border
    let fmt_c_s9 = normal
        .clone()
        .set_border_top(border_thin)
        .set_border_bottom(border_medium);
    ws.write_blank(s + 9, 2, &fmt_c_s9)?;

    // D: top+bottom+right thin border (KEIN left border!)
    let fmt_d_s9 = normal
        .clone()
        .set_border_top(border_thin)
        .set_border_bottom(border_medium)
        .set_border_right(border_thin);
    ws.write_blank(s + 9, 3, &fmt_d_s9)?;

    // E: Input-Feld, all borders
    let fmt_e_s9 = input_format
        .clone()
        .set_border_top(border_thin)
        .set_border_bottom(border_medium)
        .set_border_right(border_medium)
        .set_border_left(border_thin);
    ws.write_blank(s + 9, 4, &fmt_e_s9)?;

    // =========================================================================
    // ZEILEN 10-12 (s+10 bis s+12): Leer - keine Formatierung
    // =========================================================================

    // =========================================================================
    // ZEILE 13 (s+13): Bestätigung 1
    // =========================================================================

    write_vlookup_formula(ws, s + 13, 1, 50, &normal, language)?;

    // =========================================================================
    // ZEILE 14 (s+14): Bestätigung 2
    // =========================================================================

    write_vlookup_formula(ws, s + 14, 1, 54, &normal, language)?;

    // =========================================================================
    // ZEILEN 15-18 (s+15 bis s+18): Leer
    // =========================================================================

    // =========================================================================
    // ZEILE 19 (s+19): Unterschriften
    // =========================================================================

    // B: VLOOKUP 51, bold, top thin border
    let fmt_signature = bold.clone().set_border_top(border_thin);
    write_vlookup_formula(ws, s + 19, 1, 51, &fmt_signature, language)?;

    // C: top thin border
    let fmt_c_sig = normal.clone().set_border_top(border_thin);
    ws.write_blank(s + 19, 2, &fmt_c_sig)?;

    // D: VLOOKUP 52, bold, top thin border
    write_vlookup_formula(ws, s + 19, 3, 52, &fmt_signature, language)?;

    // E, F, G: top thin border
    let fmt_efg_sig = normal.clone().set_border_top(border_thin);
    for col in 4..=6 {
        ws.write_blank(s + 19, col, &fmt_efg_sig)?;
    }

    // =========================================================================
    // ZEILE 20 (s+20): Funktion
    // =========================================================================

    write_vlookup_formula(ws, s + 20, 3, 53, &normal, language)?;

    Ok(layout)
}

/// Schreibt die Footer Input-Werte (Bank, Kasse, Sonstiges)
pub fn write_footer_values(
    ws: &mut Worksheet,
    layout: &FooterLayout,
    styles: &ReportStyles,
    bank: Option<f64>,
    kasse: Option<f64>,
    sonstiges: Option<f64>,
) -> Result<(), XlsxError> {
    let s = layout.start_row;

    // Input-Format mit Hintergrundfarbe, Zahlenformat und Borders
    let base_input = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_unlocked()
        .set_background_color(styles.fill_input)
        .set_align(FormatAlign::Right)
        .set_align(FormatAlign::VerticalCenter)
        .set_num_format("#,##0.00");

    let border_medium = FormatBorder::Medium;
    let border_thin = FormatBorder::Thin;

    // Bank (E, Zeile 7) und Kasse (E, Zeile 8)
    let fmt_input_e = base_input
        .clone()
        .set_border_top(border_thin)
        .set_border_right(border_medium)
        .set_border_left(border_thin);

    // Sonstiges (E, Zeile 9) - mit bottom border
    let fmt_input_e_bottom = base_input
        .clone()
        .set_border_bottom(border_medium)
        .set_border_right(border_medium)
        .set_border_top(border_thin)
        .set_border_left(border_thin);

    // Bank (E, Zeile 7)
    if let Some(value) = bank {
        ws.write_number_with_format(s + 7, 4, value, &fmt_input_e)?;
    }

    // Kasse (E, Zeile 8)
    if let Some(value) = kasse {
        ws.write_number_with_format(s + 8, 4, value, &fmt_input_e)?;
    }

    // Sonstiges (E, Zeile 9)
    if let Some(value) = sonstiges {
        ws.write_number_with_format(s + 9, 4, value, &fmt_input_e_bottom)?;
    }

    Ok(())
}

/// Schreibt eine VLOOKUP-Formel mit gecachtem Text-Ergebnis
fn write_vlookup_formula(
    ws: &mut Worksheet,
    row: u32,
    col: u16,
    index: usize,
    format: &Format,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    let formula_str = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
        index
    );

    let formula = if let Some(text) = lookup_text_string(language, index) {
        Formula::new(&formula_str).set_result(text)
    } else {
        Formula::new(&formula_str)
    };

    ws.write_formula_with_format(row, col, formula, format)?;
    Ok(())
}

/// Schreibt eine VLOOKUP-Formel in einem gemergten Bereich mit gecachtem Text-Ergebnis
fn write_merged_vlookup_formula(
    ws: &mut Worksheet,
    row_start: u32,
    col_start: u16,
    row_end: u32,
    col_end: u16,
    index: usize,
    format: &Format,
    language: Option<&str>,
) -> Result<(), XlsxError> {
    let formula_str = format!(
        r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
        index
    );

    let formula = if let Some(text) = lookup_text_string(language, index) {
        Formula::new(&formula_str).set_result(text)
    } else {
        Formula::new(&formula_str)
    };

    ws.merge_range(row_start, col_start, row_end, col_end, "", format)?;
    ws.write_formula_with_format(row_start, col_start, formula, format)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_footer_layout_compute() {
        // Total in Zeile 100 (0-indexed)
        let layout = FooterLayout::compute(100);

        assert_eq!(layout.start_row, 103); // 100 + 3
        assert_eq!(layout.saldo_row, 107); // 103 + 4
        assert_eq!(layout.input_rows[0], 110); // Bank: 103 + 7
        assert_eq!(layout.input_rows[1], 111); // Kasse: 103 + 8
        assert_eq!(layout.input_rows[2], 112); // Sonstiges: 103 + 9
        assert_eq!(layout.end_row, 123); // 103 + 20
    }
}
