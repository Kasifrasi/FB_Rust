//! Header Section - Neues zellbasiertes System (Proof of Concept)
//!
//! Dieses Modul demonstriert das neue Format-System:
//! - Zellbasiert: Jede Zelle explizit mit allen Properties
//! - Fill via Variable: FillColors für gewollten Welleneffekt
//! - Merged Cells zuerst: Dann Borders (separate Ebene)
//!
//! # Beispiel-Zellen im Header
//!
//! - E2: Sprache (Input, gelb, unlocked)
//! - E3: Währung (Input, gelb, unlocked)
//! - D5: Projektnummer (Input, gelb, unlocked)
//! - D6:H7: Projekttitel (Input, gelb, unlocked, merged)

use crate::report::api::{ApiKey, ReportValues};
use crate::report::format::{BorderManager, BorderSpec, FillColors, MergedCellRegistry};
use rust_xlsxwriter::{Format, FormatAlign, Worksheet, XlsxError};

/// Schreibt Header Section mit neuem zellbasiertem System
///
/// # Workflow
///
/// 1. Merged Cells registrieren
/// 2. Borders konfigurieren (VOR dem Schreiben!)
/// 3. Zellen schreiben mit gemerged Borders
///
/// # Arguments
///
/// * `ws` - Worksheet
/// * `values` - API Values (optional)
/// * `fills` - Fill Colors (semantische Farben)
pub fn write_header_new(
    ws: &mut Worksheet,
    values: Option<&ReportValues>,
    fills: &FillColors,
) -> Result<(), XlsxError> {
    // =========================================================================
    // PHASE 1: MERGED CELLS REGISTRIEREN (ZUERST!)
    // =========================================================================

    let mut merged_cells = MergedCellRegistry::new();

    // B1:C1 - Titel "NABU-Stiftung..."
    merged_cells.register_merge(0, 1, 0, 2);

    // B2:C2 - Suffix Label
    merged_cells.register_merge(1, 1, 1, 2);

    // D6:H7 - Projekttitel (2x5 Zellen)
    merged_cells.register_merge(5, 3, 6, 7);

    // B8:C8 - Label "Projektlaufzeit von:"
    merged_cells.register_merge(7, 1, 7, 2);

    // G8:H8 - "bis"
    merged_cells.register_merge(7, 6, 7, 7);

    // B9:C9 - Label "Berichtszeitraum von:"
    merged_cells.register_merge(8, 1, 8, 2);

    // G9:H9 - "bis"
    merged_cells.register_merge(8, 6, 8, 7);

    // =========================================================================
    // PHASE 2: BORDERS KONFIGURIEREN (VOR dem Schreiben!)
    // =========================================================================

    let mut borders = BorderManager::new();
    borders.set_merged_cells(merged_cells.clone());

    // Box um Projekt-Info (D5:H7)
    borders.add_range(4, 3, 6, 7, BorderSpec::all_medium());

    // Box um Projektlaufzeit (B8:H8)
    borders.add_range(7, 1, 7, 7, BorderSpec::all_thin());

    // Box um Berichtszeitraum (B9:H9)
    borders.add_range(8, 1, 8, 7, BorderSpec::all_thin());

    // =========================================================================
    // PHASE 3: ZELLEN SCHREIBEN (Zellbasiert, mit Border-Merge)
    // =========================================================================

    // --- ZEILE 1 (Row 0) ---

    // B1 (merged B1:C1): Titel
    let mut fmt_b1 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_bold()
        .set_align(FormatAlign::Left);

    // Merge mit Border falls vorhanden
    if let Ok(Some(border_fmt)) = borders.get_border_for_cell(0, 1) {
        fmt_b1 = fmt_b1.merge(&border_fmt);
    }

    ws.merge_range(0, 1, 0, 2, "NABU-Stiftung Nationales Naturerbe", &fmt_b1)?;

    // --- ZEILE 2 (Row 1) ---

    // B2 (merged B2:C2): Suffix
    let fmt_b2 = Format::new()
        .set_font_name("Arial")
        .set_font_size(10.0)
        .set_bold()
        .set_align(FormatAlign::Left);

    let suffix_text = values
        .and_then(|v| v.get(ApiKey::Language).as_text())
        .unwrap_or("");

    ws.merge_range(1, 1, 1, 2, suffix_text, &fmt_b2)?;

    // D2: Label "Sprache:"
    ws.write_string_with_format(
        1,
        3,
        "Sprache:",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Left),
    )?;

    // E2: Sprache Input (API-Feld, gelb, unlocked)
    let lang_value = values
        .and_then(|v| v.get(ApiKey::Language).as_text())
        .unwrap_or("");

    ws.write_string_with_format(
        1,
        4,
        lang_value,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Left)
            .set_unlocked()
            .set_num_format("@"), // Text format
    )?;

    // --- ZEILE 3 (Row 2) ---

    // D3: Label "Währung:"
    ws.write_string_with_format(
        2,
        3,
        "Währung:",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Left),
    )?;

    // E3: Währung Input (API-Feld, gelb, unlocked)
    let currency_value = values
        .and_then(|v| v.get(ApiKey::Currency).as_text())
        .unwrap_or("");

    ws.write_string_with_format(
        2,
        4,
        currency_value,
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Left)
            .set_unlocked()
            .set_num_format("@"),
    )?;

    // --- ZEILE 5 (Row 4) ---

    // D5: Projektnummer Input (API-Feld, gelb, unlocked)
    let project_number = values
        .and_then(|v| v.get(ApiKey::ProjectNumber).as_text())
        .unwrap_or("");

    let fmt_d5 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Left)
            .set_unlocked()
            .set_num_format("@"),
        &borders,
        4,
        3,
    );

    ws.write_string_with_format(4, 3, project_number, &fmt_d5)?;

    // --- ZEILE 6-7 (Row 5-6, merged D6:H7) ---

    // D6:H7: Projekttitel (merged, Input, gelb, unlocked, mehrzeilig)
    let fmt_title = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Left)
            .set_text_wrap()
            .set_unlocked()
            .set_num_format("@"),
        &borders,
        5,
        3,
    );

    let project_title = values
        .and_then(|v| v.get(ApiKey::ProjectTitle).as_text())
        .unwrap_or("");

    // 1. Merge mit Format
    ws.merge_range(5, 3, 6, 7, "", &fmt_title)?;
    // 2. Write auf top-left
    ws.write_string_with_format(5, 3, project_title, &fmt_title)?;

    // --- ZEILE 8 (Row 7) ---

    // B8 (merged B8:C8): "Projektlaufzeit von:"
    let fmt_b8 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Left),
        &borders,
        7,
        1,
    );

    ws.merge_range(7, 1, 7, 2, "Projektlaufzeit von:", &fmt_b8)?;

    // D8: Label "von Datum"
    ws.write_string_with_format(
        7,
        3,
        "von Datum",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // E8: ProjectStart Input (Date, gelb, unlocked)
    let project_start = values
        .and_then(|v| v.get(ApiKey::ProjectStart).as_text())
        .unwrap_or("");

    let fmt_e8 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Center)
            .set_unlocked()
            .set_num_format("dd.mm.yyyy"),
        &borders,
        7,
        4,
    );

    ws.write_string_with_format(7, 4, project_start, &fmt_e8)?;

    // F8: "bis Datum"
    ws.write_string_with_format(
        7,
        5,
        "bis Datum",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // G8 (merged G8:H8): ProjectEnd Input
    let project_end = values
        .and_then(|v| v.get(ApiKey::ProjectEnd).as_text())
        .unwrap_or("");

    let fmt_g8 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Center)
            .set_unlocked()
            .set_num_format("dd.mm.yyyy"),
        &borders,
        7,
        6,
    );

    ws.merge_range(7, 6, 7, 7, project_end, &fmt_g8)?;

    // --- ZEILE 9 (Row 8) ---

    // B9 (merged B9:C9): "Berichtszeitraum von:"
    let fmt_b9 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_bold()
            .set_align(FormatAlign::Left),
        &borders,
        8,
        1,
    );

    ws.merge_range(8, 1, 8, 2, "Berichtszeitraum von:", &fmt_b9)?;

    // D9: "von Datum"
    ws.write_string_with_format(
        8,
        3,
        "von Datum",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // E9: ReportStart Input
    let report_start = values
        .and_then(|v| v.get(ApiKey::ReportStart).as_text())
        .unwrap_or("");

    let fmt_e9 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Center)
            .set_unlocked()
            .set_num_format("dd.mm.yyyy"),
        &borders,
        8,
        4,
    );

    ws.write_string_with_format(8, 4, report_start, &fmt_e9)?;

    // F9: "bis Datum"
    ws.write_string_with_format(
        8,
        5,
        "bis Datum",
        &Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_align(FormatAlign::Center),
    )?;

    // G9 (merged G9:H9): ReportEnd Input
    let report_end = values
        .and_then(|v| v.get(ApiKey::ReportEnd).as_text())
        .unwrap_or("");

    let fmt_g9 = with_borders(
        Format::new()
            .set_font_name("Arial")
            .set_font_size(10.0)
            .set_background_color(fills.input) // ← Fill via Variable!
            .set_align(FormatAlign::Center)
            .set_unlocked()
            .set_num_format("dd.mm.yyyy"),
        &borders,
        8,
        6,
    );

    ws.merge_range(8, 6, 8, 7, report_end, &fmt_g9)?;

    Ok(())
}

/// Merged ein Format mit Borders falls für die Zelle vorhanden
///
/// Helper-Funktion für das Baukastenprinzip: Format + Border = Finales Format
///
/// # Arguments
///
/// * `format` - Das ursprüngliche Format
/// * `borders` - BorderManager mit konfigurierten Borders
/// * `row` - Zeile
/// * `col` - Spalte
///
/// # Returns
///
/// Das gemerged Format (mit Borders falls vorhanden)
fn with_borders(format: Format, borders: &BorderManager, row: u32, col: u16) -> Format {
    if let Ok(Some(border_fmt)) = borders.get_border_for_cell(row, col) {
        format.merge(&border_fmt)
    } else {
        format
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    #[test]
    fn test_header_new_compiles() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();

        // Sollte ohne Fehler kompilieren und ausführen
        let result = write_header_new(ws, None, &fills);
        assert!(result.is_ok());
    }

    #[test]
    fn test_header_new_with_values() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let fills = FillColors::new();

        let values = ReportValues::new()
            .with_language("Deutsch")
            .with_currency("EUR")
            .with_project_number("12345")
            .with_project_title("Test Projekt");

        let result = write_header_new(ws, Some(&values), &fills);
        assert!(result.is_ok());
    }
}
