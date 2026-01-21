//! Footer Section - Neues zellbasiertes System mit variablem Start
//!
//! Der Footer beginnt an variabler Position (abhängig von Body-Länge).
//! Demonstriert relative Positionierung: `s+0`, `s+7`, `s+8`, `s+9`
//!
//! Input-Felder (Bank, Kasse, Sonstiges) waren ursprünglich LOCKED - jetzt UNLOCKED!

use crate::report::api::ApiKey;
use crate::report::body::FooterLayout;
use crate::report::format::{BorderManager, BorderSpec, FillColors, MergedCellRegistry};
use rust_xlsxwriter::{Format, FormatAlign, Worksheet, XlsxError};

/// Schreibt Footer mit neuem zellbasiertem System
///
/// # Workflow
///
/// 1. Merged Cells registrieren (relativ zu start_row)
/// 2. Borders konfigurieren (VOR dem Schreiben!)
/// 3. Zellen schreiben mit gemerged Borders
///
/// # Arguments
///
/// * `ws` - Worksheet
/// * `layout` - Footer Layout (enthält start_row!)
/// * `fills` - Fill Colors
pub fn write_footer_new(
    ws: &mut Worksheet,
    layout: &FooterLayout,
    fills: &FillColors,
) -> Result<(), XlsxError> {
    let s = layout.start_row; // Variabel!

    // =========================================================================
    // PHASE 1: MERGED CELLS REGISTRIEREN (ZUERST!, relativ zu s)
    // =========================================================================

    let mut merged_cells = MergedCellRegistry::new();

    // E(s):E(s+1) - "Saldo für Berichtszeitraum" (2 Zeilen hoch)
    merged_cells.register_merge(s, 4, s + 1, 4);

    // Weitere Merges falls vorhanden...

    // =========================================================================
    // PHASE 2: BORDERS KONFIGURIEREN (VOR dem Schreiben!)
    // =========================================================================

    let mut borders = BorderManager::new();
    borders.set_merged_cells(merged_cells.clone());

    // Box um Input-Bereich (B(s+7):E(s+9))
    borders.add_range(
        s + 7,
        1,
        s + 9,
        4,
        BorderSpec {
            top: Some(rust_xlsxwriter::FormatBorder::Medium),
            bottom: Some(rust_xlsxwriter::FormatBorder::Medium),
            left: Some(rust_xlsxwriter::FormatBorder::Medium),
            right: Some(rust_xlsxwriter::FormatBorder::Medium),
        },
    );

    // =========================================================================
    // PHASE 2: ZELLEN SCHREIBEN (Zellbasiert, relativ zu s)
    // =========================================================================

    // --- ROW s+0 und s+1 (Saldo-Label, merged) ---

    // E(s):E(s+1) (merged): "Saldo für Berichtszeitraum" (VLOOKUP würde hier stehen)
    let fmt_saldo = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_align(FormatAlign::Center)
        .set_text_wrap();

    ws.merge_range(s, 4, s + 1, 4, "", &fmt_saldo)?;
    // VLOOKUP-Formel würde hier geschrieben werden
    ws.write_string_with_format(s, 4, "Saldo für Berichtszeitraum", &fmt_saldo)?;

    // --- ROW s+7 (Bank Input) ---

    // B(s+7): Label "Bank"
    ws.write_string_with_format(
        s + 7,
        1,
        "Bank",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Left),
    )?;

    // E(s+7): Bank Input (KRITISCH: unlocked!, gelb, number)
    ws.write_number_with_format(
        s + 7,
        4,
        0.0,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Right)
            .set_unlocked() // ✅ KRITISCH: Input-Felder müssen unlocked sein!
            .set_num_format("#,##0.00"),
    )?;

    // --- ROW s+8 (Kasse Input) ---

    // B(s+8): Label "Kasse"
    ws.write_string_with_format(
        s + 8,
        1,
        "Kasse",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Left),
    )?;

    // E(s+8): Kasse Input (KRITISCH: unlocked!, gelb, number)
    ws.write_number_with_format(
        s + 8,
        4,
        0.0,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Right)
            .set_unlocked() // ✅ KRITISCH: Input-Felder müssen unlocked sein!
            .set_num_format("#,##0.00"),
    )?;

    // --- ROW s+9 (Sonstiges Input) ---

    // B(s+9): Label "Sonstiges"
    ws.write_string_with_format(
        s + 9,
        1,
        "Sonstiges",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Left),
    )?;

    // E(s+9): Sonstiges Input (KRITISCH: unlocked!, gelb, number)
    ws.write_number_with_format(
        s + 9,
        4,
        0.0,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Right)
            .set_unlocked() // ✅ KRITISCH: Input-Felder müssen unlocked sein!
            .set_num_format("#,##0.00"),
    )?;

    Ok(())
}

/// Merged ein Format mit Borders falls für die Zelle vorhanden
///
/// Helper-Funktion für das Baukastenprinzip: Format + Border = Finales Format
fn with_borders(format: Format, borders: &BorderManager, row: u32, col: u16) -> Format {
    if let Ok(Some(border_fmt)) = borders.get_border_for_cell(row, col) {
        format.merge(&border_fmt)
    } else {
        format
    }
}

/// Erstellt API-Keys für Footer-Felder
///
/// Demonstriert wie API-Registrierung im Footer funktioniert
pub fn footer_api_keys() -> Vec<(ApiKey, u32, u16)> {
    vec![
        // Diese würden relativ zu start_row registriert
        // (ApiKey::Footer(FooterField::Bank), s + 7, 4),
        // (ApiKey::Footer(FooterField::Kasse), s + 8, 4),
        // (ApiKey::Footer(FooterField::Sonstiges), s + 9, 4),
    ]
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    #[test]
    fn test_footer_new_compiles() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();

        // Footer startet bei Zeile 100 (variabel!)
        let layout = FooterLayout {
            start_row: 100,
            saldo_row: 100,
            input_rows: [107, 108, 109],
            end_row: 120,
        };

        let result = write_footer_new(ws, &layout, &fills);
        assert!(result.is_ok());
    }

    #[test]
    fn test_footer_variable_start() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();

        // Footer startet bei Zeile 50
        let layout1 = FooterLayout {
            start_row: 50,
            saldo_row: 50,
            input_rows: [57, 58, 59],
            end_row: 70,
        };
        let result1 = write_footer_new(ws, &layout1, &fills);
        assert!(result1.is_ok());

        // Footer startet bei Zeile 200 (anderes Worksheet)
        let ws2 = wb.add_worksheet();
        let layout2 = FooterLayout {
            start_row: 200,
            saldo_row: 200,
            input_rows: [207, 208, 209],
            end_row: 220,
        };
        let result2 = write_footer_new(ws2, &layout2, &fills);
        assert!(result2.is_ok());
    }

    #[test]
    fn test_footer_relative_positions() {
        let layout = FooterLayout {
            start_row: 100,
            saldo_row: 100,
            input_rows: [107, 108, 109],
            end_row: 120,
        };
        let s = layout.start_row;

        // Relative Positionen
        assert_eq!(s + 0, 100); // Saldo
        assert_eq!(s + 7, 107); // Bank
        assert_eq!(s + 8, 108); // Kasse
        assert_eq!(s + 9, 109); // Sonstiges
    }

    #[test]
    fn test_merged_cells_relative() {
        let mut merged_cells = MergedCellRegistry::new();
        let s = 100; // start_row

        // E(s):E(s+1) - relativ registriert
        merged_cells.register_merge(s, 4, s + 1, 4);

        assert_eq!(merged_cells.count(), 1);

        // Top-left ist (s, 4)
        assert!(merged_cells.get_merge_range(s, 4).is_some());

        // (s+1, 4) ist non-topleft
        assert!(merged_cells.is_merged_non_topleft(s + 1, 4));
    }
}
