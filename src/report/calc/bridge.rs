//! IronCalc-Bridge: Dynamische Formeln, Input-Population, Evaluation
//!
//! Ersetzt `CellRegistry` + `evaluate_all_cells()` + `write_cells_from_registry()`.

use super::template::ModelTemplate;
use crate::report::api::{ApiKey, CellValue, FooterField, PositionField, ReportValues};
use crate::report::body::{
    BodyLayout, CategoryLayout, CategoryMode, FooterLayout, TOTAL_LABEL_INDEX,
};
use crate::report::core::CellAddr;
use ironcalc_base::cell::CellValue as IcCellValue;
use ironcalc_base::Model;

/// Sheet-Index für den Bericht
const SHEET: u32 = 0;

/// Bridge zwischen ReportValues und IronCalc Model.
///
/// Verantwortlich für:
/// 1. Dynamische Formeln hinzufügen (Body, Footer)
/// 2. Input-Werte setzen
/// 3. Evaluation triggern
/// 4. Ergebnisse auslesen
pub(crate) struct CalcBridge {
    model: Model<'static>,
    /// Alle Zellen mit Formeln (für write_cells_from_bridge)
    formula_cells: Vec<CellAddr>,
    /// Alle Zellen mit Input-Werten (für write_cells_from_bridge)
    input_cells: Vec<CellAddr>,
    /// Hyperlink-Zellen — IronCalc evaluiert den VLOOKUP, Writer schreibt nativen Hyperlink
    hyperlink_cells: Vec<CellAddr>,
}

impl CalcBridge {
    /// Erstellt eine Bridge aus einem Template.
    ///
    /// Übernimmt die statischen Formel-Adressen aus dem Template,
    /// damit `write_cells_from_bridge()` sie in die Excel-Datei schreibt.
    pub(crate) fn from_template(template: &ModelTemplate) -> Self {
        Self {
            model: template.instantiate(),
            formula_cells: template.static_formula_cells().to_vec(),
            input_cells: Vec::new(),
            hyperlink_cells: template.hyperlink_cells().to_vec(),  // Nur Adressen, Writer nutzt write_url
        }
    }

    // ========================================================================
    // Dynamische Body-Formeln
    // ========================================================================

    /// Fügt dynamische Body-Formeln hinzu basierend auf Layout.
    ///
    /// Ersetzt: registration.rs `register_body_formulas()`
    pub(crate) fn register_body_formulas(&mut self, layout: &BodyLayout) {
        // 1. VLOOKUP-Formeln für Labels
        for cat in &layout.categories {
            match &cat.mode {
                CategoryMode::HeaderInput { row } => {
                    // C-Spalte: VLOOKUP für Kategorie-Label
                    self.add_vlookup(*row, 2, cat.meta.label_index);
                }
                CategoryMode::WithPositions {
                    header_row,
                    footer_row,
                    ..
                } => {
                    // Header: C-Spalte
                    self.add_vlookup(*header_row, 2, cat.meta.label_index);
                    // Footer: B-Spalte
                    self.add_vlookup(*footer_row, 1, cat.meta.sum_label_index);
                }
            }
        }

        // Total: B-Spalte
        self.add_vlookup(layout.total_row, 1, TOTAL_LABEL_INDEX);

        // 2. Ratio-Formeln (G-Spalte)
        for row in layout.ratio_rows() {
            self.add_ratio(row);
        }

        // 3. SUM-Formeln für Footer-Zeilen (D, E, F)
        for cat in &layout.categories {
            if let CategoryMode::WithPositions {
                positions,
                footer_row,
                ..
            } = &cat.mode
            {
                for col in [3u16, 4, 5] {
                    self.add_sum_range(*footer_row, col, positions.start_row, positions.end_row);
                }
            }
        }

        // 4. SUM-Formeln für Total-Zeile (D, E, F)
        let sum_rows: Vec<u32> = layout
            .categories
            .iter()
            .map(|c: &CategoryLayout| c.sum_row())
            .collect();
        for col in [3u16, 4, 5] {
            self.add_sum_cells(layout.total_row, col, &sum_rows);
        }
    }

    // ========================================================================
    // Dynamische Footer-Formeln
    // ========================================================================

    /// Fügt dynamische Footer-Formeln hinzu.
    ///
    /// Ersetzt: registration.rs `register_footer_formulas()`
    pub(crate) fn register_footer_formulas(
        &mut self,
        footer_layout: &FooterLayout,
        income_row: u32,
    ) {
        let s = footer_layout.start_row;
        let total_row = s - 3; // Footer startet 3 Zeilen nach Total

        // VLOOKUP-Labels
        self.add_vlookup(s, 4, 44); // E(s): "Saldo für den Berichtszeitraum"
        self.add_vlookup(s + 1, 1, 43); // B(s+1): "ABSCHLUSS"
        self.add_vlookup(s + 2, 4, 10); // E(s+2): Währung
        self.add_vlookup(s + 4, 1, 45); // B(s+4): "Saldo..."
        self.add_vlookup(s + 6, 1, 46); // B(s+6): "Saldenabstimmung:"
        self.add_vlookup(s + 7, 1, 47); // B(s+7): "Bank"
        self.add_vlookup(s + 8, 1, 48); // B(s+8): "Kasse"
        self.add_vlookup(s + 9, 1, 49); // B(s+9): "Sonstiges"
        self.add_vlookup(s + 13, 1, 50); // B(s+13): Bestätigung 1
        self.add_vlookup(s + 14, 1, 54); // B(s+14): Bestätigung 2
        self.add_vlookup(s + 19, 1, 51); // B(s+19): "Ort, Datum..."
        self.add_vlookup(s + 19, 3, 52); // D(s+19): "Unterschrift..."
        self.add_vlookup(s + 20, 3, 53); // D(s+20): "Funktion..."

        // Check-Formeln
        // D(s+4): =IF(ROUND(E_saldo,2)=(ROUND(F_income-F_total,2)),"✓","")
        let saldo_row_excel = s + 4 + 1;
        let income_row_excel = income_row + 1;
        let total_row_excel = total_row + 1;
        self.add_formula(
            s + 4,
            3,
            format!(
                r#"=IF(ROUND(E{},2)=(ROUND(F{}-F{},2)),"✓","")"#,
                saldo_row_excel, income_row_excel, total_row_excel
            ),
        );

        // E(s+4): Differenz = E_income - E_total
        self.add_formula(
            s + 4,
            4,
            format!("=E{}-E{}", income_row_excel, total_row_excel),
        );

        // E(s+6): =IF(E_saldo=SUM(E_bank:E_sonstiges),"OK","")
        let bank_row_excel = footer_layout.input_rows[0] + 1;
        let sonstiges_row_excel = footer_layout.input_rows[2] + 1;
        self.add_formula(
            s + 6,
            4,
            format!(
                r#"=IF(E{}=SUM(E{}:E{}),"OK","")"#,
                saldo_row_excel, bank_row_excel, sonstiges_row_excel
            ),
        );
    }

    // ========================================================================
    // Input-Population
    // ========================================================================

    /// Setzt alle Input-Werte aus ReportValues.
    ///
    /// Ersetzt: evaluate_all_cells() Phase 1 (API-Werte eintragen)
    pub(crate) fn populate(
        &mut self,
        values: &ReportValues,
        layout: &BodyLayout,
        footer_layout: &FooterLayout,
    ) {
        // Statische API-Keys (Header, Table, Panel)
        for key in ApiKey::all_static_keys() {
            let value = values.get(key);
            if let Some(addr) = key.static_addr() {
                self.set_cell_value(addr, value);
            }
        }

        // Dynamische Position-Keys (Body)
        for cat in &layout.categories {
            match &cat.mode {
                CategoryMode::HeaderInput { row } => {
                    for field in PositionField::header_input_fields() {
                        let key = ApiKey::Position {
                            category: cat.meta.num,
                            position: 0,
                            field,
                        };
                        let addr = CellAddr::new(*row, field.col());
                        self.set_cell_value(addr, values.get(key));
                    }
                }
                CategoryMode::WithPositions { positions, .. } => {
                    for i in 0..positions.count {
                        let row = positions.start_row + i as u32;
                        let pos = i + 1;
                        for field in PositionField::all() {
                            let key = ApiKey::Position {
                                category: cat.meta.num,
                                position: pos,
                                field,
                            };
                            let addr = CellAddr::new(row, field.col());
                            self.set_cell_value(addr, values.get(key));
                        }
                    }
                }
            }
        }

        // Footer-Keys (Bank, Kasse, Sonstiges)
        for field in FooterField::all() {
            let key = ApiKey::Footer(field);
            let addr = CellAddr::new(footer_layout.input_rows[field.index()], field.col());
            self.set_cell_value(addr, values.get(key));
        }
    }

    /// Evaluiert alle Formeln.
    pub(crate) fn evaluate(&mut self) {
        self.model.evaluate();
    }

    // ========================================================================
    // Ergebnis-Zugriff
    // ========================================================================

    /// Liest einen berechneten Zellwert (konvertiert IronCalc → CellValue).
    pub(crate) fn get_value(&self, row: u32, col: u16) -> CellValue {
        let ic_val = self
            .model
            .get_cell_value_by_index(SHEET, row as i32 + 1, col as i32 + 1)
            .unwrap_or(IcCellValue::None);
        convert_cell_value(ic_val)
    }

    /// Liest einen Formel-String (None wenn keine Formel).
    pub(crate) fn get_formula(&self, row: u32, col: u16) -> Option<String> {
        self.model
            .get_cell_formula(SHEET, row as i32 + 1, col as i32 + 1)
            .ok()
            .flatten()
    }

    /// Gibt alle Zellen mit Formeln zurück.
    pub(crate) fn formula_cells(&self) -> &[CellAddr] {
        &self.formula_cells
    }

    /// Gibt alle Zellen mit Input-Werten zurück.
    pub(crate) fn input_cells(&self) -> &[CellAddr] {
        &self.input_cells
    }

    /// Gibt Hyperlink-Zellen zurück.
    ///
    /// IronCalc evaluiert den VLOOKUP (→ URL-String),
    /// der Writer schreibt einen nativen Hyperlink via `write_url_with_format()`.
    pub(crate) fn hyperlink_cells(&self) -> &[CellAddr] {
        &self.hyperlink_cells
    }

    // ========================================================================
    // Private Helpers
    // ========================================================================

    /// Setzt eine VLOOKUP-Formel und trackt die Position.
    fn add_vlookup(&mut self, row: u32, col: u16, index: usize) {
        let formula = format!(
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
            index
        );
        self.add_formula(row, col, formula);
    }

    /// Setzt eine Ratio-Formel: =IFERROR(F{row}/D{row},0)
    fn add_ratio(&mut self, row: u32) {
        let excel_row = row + 1;
        let formula = format!("=IFERROR(F{}/D{},0)", excel_row, excel_row);
        self.add_formula(row, 6, formula);
    }

    /// Setzt eine SUM-Range-Formel: =SUM(D{start}:D{end})
    fn add_sum_range(&mut self, target_row: u32, col: u16, start_row: u32, end_row: u32) {
        let col_letter = CellAddr::col_to_letter(col);
        let formula = format!(
            "=SUM({}{}:{}{})",
            col_letter,
            start_row + 1,
            col_letter,
            end_row + 1
        );
        self.add_formula(target_row, col, formula);
    }

    /// Setzt eine SUM-Formel über einzelne Zellen: =SUM(D47+D69+...)
    fn add_sum_cells(&mut self, target_row: u32, col: u16, source_rows: &[u32]) {
        let col_letter = CellAddr::col_to_letter(col);
        let refs: Vec<String> = source_rows
            .iter()
            .map(|r| format!("{}{}", col_letter, r + 1))
            .collect();
        let formula = format!("=SUM({})", refs.join("+"));
        self.add_formula(target_row, col, formula);
    }

    /// Low-level: Setzt eine Formel und trackt die Position.
    fn add_formula(&mut self, row: u32, col: u16, formula: String) {
        self.model
            .update_cell_with_formula(SHEET, row as i32 + 1, col as i32 + 1, formula)
            .unwrap_or_else(|e| panic!("Failed to set formula at ({}, {}): {}", row, col, e));
        self.formula_cells.push(CellAddr::new(row, col));
    }

    /// Setzt einen CellValue in IronCalc und trackt die Position.
    fn set_cell_value(&mut self, addr: CellAddr, value: &CellValue) {
        let ic_row = addr.row as i32 + 1;
        let ic_col = addr.col as i32 + 1;

        match value {
            CellValue::Empty => {
                // Leere Zellen brauchen wir in IronCalc nicht zu setzen,
                // aber wir tracken sie für die Excel-Ausgabe (Blanks mit Format).
                self.input_cells.push(addr);
            }
            CellValue::Text(s) | CellValue::Date(s) => {
                self.model
                    .set_user_input(SHEET, ic_row, ic_col, s.clone())
                    .unwrap_or_else(|e| {
                        panic!("Failed to set text at ({}, {}): {}", addr.row, addr.col, e)
                    });
                self.input_cells.push(addr);
            }
            CellValue::Number(n) => {
                // Zahlen als String setzen (IronCalc parsed den Typ automatisch)
                self.model
                    .set_user_input(SHEET, ic_row, ic_col, n.to_string())
                    .unwrap_or_else(|e| {
                        panic!("Failed to set number at ({}, {}): {}", addr.row, addr.col, e)
                    });
                self.input_cells.push(addr);
            }
        }
    }
}

/// Konvertiert IronCalc CellValue → Project CellValue
fn convert_cell_value(ic: IcCellValue) -> CellValue {
    match ic {
        IcCellValue::None => CellValue::Empty,
        IcCellValue::String(s) => {
            if s.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Text(s)
            }
        }
        IcCellValue::Number(n) => CellValue::Number(n),
        IcCellValue::Boolean(b) => CellValue::Text(if b { "TRUE" } else { "FALSE" }.to_string()),
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::report::body::config::BodyConfig;

    fn create_template() -> ModelTemplate {
        ModelTemplate::new()
    }

    #[test]
    fn test_bridge_basic_flow() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0)
            .with_positions(3, 0)
            .with_positions(4, 0)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        // Werte setzen
        let values = ReportValues::new()
            .with_language("deutsch")
            .with_project_number("2025-001")
            .with_project_title("Testprojekt");

        bridge.populate(&values, &layout, &footer_layout);
        bridge.evaluate();

        // B1 sollte den VLOOKUP-Titel haben
        let b1 = bridge.get_value(0, 1);
        assert_eq!(
            b1,
            CellValue::Text("F I N A N Z B E R I C H T".to_string())
        );

        // D5 sollte die Projektnummer sein
        let d5 = bridge.get_value(4, 3);
        assert_eq!(d5, CellValue::Text("2025-001".to_string()));
    }

    #[test]
    fn test_bridge_body_formulas() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new()
            .with_positions(1, 3)
            .with_positions(2, 0)
            .with_positions(3, 0)
            .with_positions(4, 0)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        // Positionen für Kategorie 1 setzen
        let mut values = ReportValues::new().with_language("deutsch");
        values.set_position(1, 1, PositionField::Approved, 1000.0);
        values.set_position(1, 1, PositionField::IncomeTotal, 500.0);
        values.set_position(1, 2, PositionField::Approved, 2000.0);
        values.set_position(1, 2, PositionField::IncomeTotal, 1500.0);
        values.set_position(1, 3, PositionField::Approved, 3000.0);
        values.set_position(1, 3, PositionField::IncomeTotal, 2000.0);

        bridge.populate(&values, &layout, &footer_layout);
        bridge.evaluate();

        // Footer-Zeile für Kategorie 1: SUM(D27:D29) = 1000 + 2000 + 3000 = 6000
        let cat1_footer_row = layout.category(1).unwrap().footer_row().unwrap();
        let d_sum = bridge.get_value(cat1_footer_row, 3);
        assert_eq!(d_sum, CellValue::Number(6000.0));

        // F-Sum: 500 + 1500 + 2000 = 4000
        let f_sum = bridge.get_value(cat1_footer_row, 5);
        assert_eq!(f_sum, CellValue::Number(4000.0));

        // Ratio (G): IFERROR(F/D, 0) = 4000/6000 ≈ 0.6667
        let ratio = bridge.get_value(cat1_footer_row, 6);
        match ratio {
            CellValue::Number(n) => assert!((n - 4000.0 / 6000.0).abs() < 0.0001),
            other => panic!("Expected Number, got {:?}", other),
        }
    }

    #[test]
    fn test_bridge_formula_readback() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new().with_positions(1, 2);
        let layout = BodyLayout::compute(&config);

        bridge.register_body_formulas(&layout);

        // VLOOKUP-Formel für Kategorie 1 Header (C27)
        let formula = bridge.get_formula(26, 2);
        assert!(formula.is_some(), "Should have formula at C27");
        let f = formula.unwrap();
        assert!(f.contains("VLOOKUP"), "Formula should contain VLOOKUP: {}", f);
    }

    #[test]
    fn test_bridge_footer_check() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new()
            .with_positions(1, 1)
            .with_positions(2, 0)
            .with_positions(3, 0)
            .with_positions(4, 0)
            .with_positions(5, 0)
            .with_positions(6, 0)
            .with_positions(7, 0)
            .with_positions(8, 0);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        // Einnahmen-Tabelle: Gesamteinnahmen (F20)
        let mut values = ReportValues::new().with_language("deutsch");
        values.set(ApiKey::IncomeTotal(0), 1000.0); // F15

        // Body: Kategorie 1, Position 1 Ausgaben
        values.set_position(1, 1, PositionField::IncomeTotal, 800.0);

        // Footer: Saldenabstimmung
        values
            .set_footer_bank(200.0)
            .set_footer_kasse(0.0)
            .set_footer_sonstiges(0.0);

        bridge.populate(&values, &layout, &footer_layout);
        bridge.evaluate();

        // Die Differenz E(s+4) = E_income(20) - E_total
        let diff = bridge.get_value(footer_layout.start_row + 4, 4);
        assert!(
            matches!(diff, CellValue::Number(_)),
            "Diff should be a number, got {:?}",
            diff
        );
    }

    #[test]
    fn test_bridge_cell_tracking() {
        let template = create_template();
        let mut bridge = CalcBridge::from_template(&template);

        let config = BodyConfig::new().with_positions(1, 1);
        let layout = BodyLayout::compute(&config);
        let footer_layout = FooterLayout::compute(layout.total_row);

        bridge.register_body_formulas(&layout);
        bridge.register_footer_formulas(&footer_layout, 19);

        assert!(
            !bridge.formula_cells().is_empty(),
            "Should have formula cells after registration"
        );

        let values = ReportValues::new().with_language("deutsch");
        bridge.populate(&values, &layout, &footer_layout);

        assert!(
            !bridge.input_cells().is_empty(),
            "Should have input cells after populate"
        );
    }
}
