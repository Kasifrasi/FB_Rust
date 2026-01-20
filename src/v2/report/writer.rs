//! Writer V2 - Registry-basierter Finanzbericht Generator
//!
//! Nutzt die CellRegistry für:
//! - Zentrale Zell-Verwaltung
//! - Automatische Auswertungsreihenfolge
//! - Unterstützung dynamischer Bereiche

use super::body::{
    register_body_formulas, register_footer_formulas, write_body_structure_with_values,
    write_footer, write_footer_values, BodyConfig, BodyLayout, BodyResult, CategoryMode,
    FooterLayout,
};
use super::definitions::build_registry;
use super::formats::{
    build_format_matrix, extend_format_matrix_with_body, FormatMatrix, ReportStyles, SectionStyles,
};
use super::layout;
use super::registry::{CellAddr, CellKind, CellRegistry, EvalContext};
use super::sections::{
    write_header_section, write_panel_section, write_prebody_section,
    write_prebody_section_unified, write_table_section,
};
use super::values::{CellValue, ReportValues};
use rust_xlsxwriter::{Format, Formula, Worksheet, XlsxError};
use std::collections::HashMap;

// ============================================================================
// Report Writer V2
// ============================================================================

/// Schreibt den kompletten Finanzbericht (Registry-basiert)
pub fn write_report_v2(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    values: &ReportValues,
) -> Result<(), XlsxError> {
    // Standardwert für Formel-Ergebnisse auf "" setzen (statt 0)
    ws.set_formula_result_default("");

    // 1. Registry erstellen
    let registry = build_registry().map_err(|e| {
        // Konvertiere RegistryError zu XlsxError
        XlsxError::ParameterError(format!("Registry error: {}", e))
    })?;

    // 2. Alle Zellen evaluieren
    let computed = evaluate_all_cells(&registry, values);

    // 3. Styles und FormatMatrix erstellen
    let sec = SectionStyles::new(styles);
    let fmt = build_format_matrix(styles, &sec);

    // 4. Statische Sections schreiben (Layout, Merges, etc.)
    let lang_val = values.language().unwrap_or("");
    write_header_section(ws, &fmt, suffix, lang_val)?;
    write_table_section(ws, &fmt)?;
    write_panel_section(ws, &fmt, values)?;

    // 5. Alle Zellen aus Registry schreiben
    write_cells_from_registry(ws, &registry, &computed, &fmt)?;

    // 6. Freeze Pane
    layout::setup_freeze_panes(ws, 9)?;

    Ok(())
}

/// Schreibt den kompletten Finanzbericht MIT dynamischem Body-Bereich
pub fn write_report_v2_with_body(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
) -> Result<BodyResult, XlsxError> {
    // Standardwert für Formel-Ergebnisse auf "" setzen (statt 0)
    ws.set_formula_result_default("");

    // 1. Registry erstellen (für statischen Bereich)
    let registry = build_registry()
        .map_err(|e| XlsxError::ParameterError(format!("Registry error: {}", e)))?;

    // 2. Body-Layout berechnen
    let body_layout = BodyLayout::compute(body_config);

    // 3. Alle statischen Zellen evaluieren
    let computed = evaluate_all_cells(&registry, values);

    // 4. FormatMatrix erstellen (statisch + body)
    let sec = SectionStyles::new(styles);
    let mut fmt = build_format_matrix(styles, &sec);
    extend_format_matrix_with_body(&mut fmt, styles, &body_layout);

    // 5. Statische Sections schreiben
    let lang_val = values.language().unwrap_or("");
    write_header_section(ws, &fmt, suffix, lang_val)?;
    write_table_section(ws, &fmt)?;
    write_panel_section(ws, &fmt, values)?;
    write_prebody_section(ws, styles, values.language())?;

    // 6. Statische Zellen aus Registry schreiben
    write_cells_from_registry(ws, &registry, &computed, &fmt)?;

    // 7. Dynamischen Body schreiben (mit API-Werten)
    let body_result = write_body_structure_with_values(ws, &fmt, body_config, Some(values))?;

    // 8. Footer schreiben (3 Zeilen nach Total)
    // income_row = 19 (0-indexed, Zeile 20 in Excel)
    let income_row = 19u32;

    // E20 und F20 (Einnahmen-Summe) aus computed holen für Check-Formel Evaluierung
    let e_income = computed
        .get(&CellAddr::new(19, 4))
        .and_then(|v| v.as_number());
    let f_income = computed
        .get(&CellAddr::new(19, 5))
        .and_then(|v| v.as_number());

    let footer_layout = write_footer(
        ws,
        styles,
        body_result.layout.total_row,
        income_row,
        values.language(),
        e_income,
        body_result.e_total,
        f_income,
        body_result.f_total,
        values.footer_bank(),
        values.footer_kasse(),
        values.footer_sonstiges(),
    )?;

    // 9. Footer-Werte schreiben (Bank, Kasse, Sonstiges)
    write_footer_values(
        ws,
        &footer_layout,
        styles,
        values.footer_bank(),
        values.footer_kasse(),
        values.footer_sonstiges(),
    )?;

    // 10. Freeze Pane
    layout::setup_freeze_panes(ws, 9)?;

    Ok(body_result)
}

/// Schreibt den kompletten Finanzbericht mit einheitlicher Registry-basierter Evaluierung
///
/// Diese Funktion nutzt die einheitliche CellRegistry für ALLE Bereiche:
/// - Statische Zellen (A1:V25)
/// - Body-Zellen (dynamisch basierend auf BodyConfig)
/// - Prebody-Formeln
/// - Footer-Formeln
///
/// Vorteile:
/// - Einheitliche Evaluierungsreihenfolge durch topologische Sortierung
/// - Einheitliches Caching in `computed` HashMap
/// - Konsistente Formel-Registrierung für alle Bereiche
pub fn write_report_v2_with_body_unified(
    ws: &mut Worksheet,
    styles: &ReportStyles,
    suffix: &str,
    values: &ReportValues,
    body_config: &BodyConfig,
) -> Result<BodyResult, XlsxError> {
    // Standardwert für Formel-Ergebnisse auf "" setzen (statt 0)
    ws.set_formula_result_default("");

    // 1. Registry erstellen (für statischen Bereich)
    let mut registry = build_registry()
        .map_err(|e| XlsxError::ParameterError(format!("Registry error: {}", e)))?;

    // 2. Body-Layout berechnen
    let body_layout = BodyLayout::compute(body_config);

    // 3. Body-Formeln mit der Registry registrieren
    register_body_formulas(&mut registry, &body_layout).map_err(|e| {
        XlsxError::ParameterError(format!("Body formula registration error: {}", e))
    })?;

    // 4. Footer-Layout berechnen und Formeln registrieren
    let footer_layout = FooterLayout::compute(body_layout.total_row);
    let income_row = 19u32; // Einnahmen-Zeile (0-indexiert)
    register_footer_formulas(&mut registry, &footer_layout, income_row).map_err(|e| {
        XlsxError::ParameterError(format!("Footer formula registration error: {}", e))
    })?;

    // 5. ALLE Zellen evaluieren (statisch + Body + Footer, einheitlich!)
    let computed = evaluate_all_cells(&registry, values);

    // 6. FormatMatrix erstellen (statisch + body)
    let sec = SectionStyles::new(styles);
    let mut fmt = build_format_matrix(styles, &sec);
    extend_format_matrix_with_body(&mut fmt, styles, &body_layout);

    // 7. Statische Sections schreiben
    let lang_val = values.language().unwrap_or("");
    write_header_section(ws, &fmt, suffix, lang_val)?;
    write_table_section(ws, &fmt)?;
    write_panel_section(ws, &fmt, values)?;
    // Prebody: nur Layout, Formeln kommen aus Registry
    write_prebody_section_unified(ws, styles)?;

    // 8. ALLE Zellen aus Registry schreiben (statisch + Body + Footer, einheitlich!)
    write_cells_from_registry(ws, &registry, &computed, &fmt)?;

    // 9. Body-Struktur schreiben (nur Layout: Nummern, Merges, Blanks)
    // API-Werte und Formeln wurden bereits in Schritt 8 geschrieben
    let body_result = write_body_structure_unified(ws, &fmt, &body_layout, &computed)?;

    // 10. Footer schreiben (Layout + Werte aus Registry)
    // Die Formeln wurden bereits in Schritt 8 geschrieben
    let e_income = computed
        .get(&CellAddr::new(19, 4))
        .and_then(|v| v.as_number());
    let f_income = computed
        .get(&CellAddr::new(19, 5))
        .and_then(|v| v.as_number());

    // Alte write_footer Funktion für Layout (wird später refaktoriert)
    // Die Formel-Evaluierung erfolgt jetzt zentral über die Registry
    write_footer(
        ws,
        styles,
        body_result.layout.total_row,
        income_row,
        values.language(),
        e_income,
        body_result.e_total,
        f_income,
        body_result.f_total,
        values.footer_bank(),
        values.footer_kasse(),
        values.footer_sonstiges(),
    )?;

    write_footer_values(
        ws,
        &footer_layout,
        styles,
        values.footer_bank(),
        values.footer_kasse(),
        values.footer_sonstiges(),
    )?;

    // 10. Freeze Pane
    layout::setup_freeze_panes(ws, 9)?;

    Ok(body_result)
}

/// Schreibt die Body-Struktur (nur Layout-Elemente, Werte kommen aus Registry)
fn write_body_structure_unified(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    layout: &BodyLayout,
    computed: &HashMap<CellAddr, CellValue>,
) -> Result<BodyResult, XlsxError> {
    // Kategorien schreiben
    for cat in &layout.categories {
        write_category_unified(ws, fmt, cat, computed)?;
    }

    // Total-Zeile: Nummer schreiben (Formeln kommen aus Registry)
    // B:C ist gemerged mit VLOOKUP-Formel
    let total_row = layout.total_row;
    if let Some(format) = fmt.get(total_row, 1) {
        ws.merge_range(total_row, 1, total_row, 2, "", format)?;
    }
    // H: Blank
    if let Some(format) = fmt.get(total_row, 7) {
        ws.write_blank(total_row, 7, format)?;
    }

    // E-Total und F-Total aus computed holen
    let e_total = computed
        .get(&CellAddr::new(total_row, 4))
        .and_then(|v| v.as_number());
    let f_total = computed
        .get(&CellAddr::new(total_row, 5))
        .and_then(|v| v.as_number());

    Ok(BodyResult {
        layout: layout.clone(),
        last_row: layout.last_row,
        total_row: layout.total_row,
        e_total,
        f_total,
    })
}

/// Schreibt eine Kategorie (nur Layout, Werte aus Registry)
fn write_category_unified(
    ws: &mut Worksheet,
    fmt: &FormatMatrix,
    cat: &super::body::CategoryLayout,
    computed: &HashMap<CellAddr, CellValue>,
) -> Result<(), XlsxError> {
    match &cat.mode {
        CategoryMode::HeaderInput { row } => {
            // B: Kategorie-Nummer
            if let Some(format) = fmt.get(*row, 1) {
                ws.write_string_with_format(*row, 1, &format!("{}.", cat.meta.num), format)?;
            }
            // H: Blank (falls nicht API)
            if let Some(format) = fmt.get(*row, 7) {
                if computed
                    .get(&CellAddr::new(*row, 7))
                    .map_or(true, |v| v.is_empty())
                {
                    ws.write_blank(*row, 7, format)?;
                }
            }
        }
        CategoryMode::WithPositions {
            header_row,
            positions,
            footer_row,
        } => {
            // Header-Zeile: B = Nummer
            if let Some(format) = fmt.get(*header_row, 1) {
                ws.write_string_with_format(*header_row, 1, &format!("{}.", cat.meta.num), format)?;
            }
            // Header D-H: Blanks
            for col in 3..=7 {
                if let Some(format) = fmt.get(*header_row, col) {
                    ws.write_blank(*header_row, col, format)?;
                }
            }

            // Positions-Zeilen: B = Nummer
            for i in 0..positions.count {
                let row = positions.start_row + i as u32;
                let pos_num = i + 1;
                if let Some(format) = fmt.get(row, 1) {
                    ws.write_string_with_format(
                        row,
                        1,
                        &format!("{}.{}", cat.meta.num, pos_num),
                        format,
                    )?;
                }
            }

            // Footer-Zeile: B:C merged
            if let Some(format) = fmt.get(*footer_row, 1) {
                ws.merge_range(*footer_row, 1, *footer_row, 2, "", format)?;
            }
            // Footer H: Blank
            if let Some(format) = fmt.get(*footer_row, 7) {
                ws.write_blank(*footer_row, 7, format)?;
            }
        }
    }
    Ok(())
}

/// Evaluiert alle Zellen und gibt die berechneten Werte zurück
fn evaluate_all_cells(
    registry: &CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    values: &ReportValues,
) -> HashMap<CellAddr, CellValue> {
    let mut computed: HashMap<CellAddr, CellValue> = HashMap::new();

    // 1. API-Werte eintragen
    for addr in registry.api_cells() {
        if let CellKind::Api(api) = registry.get(*addr) {
            let value = get_api_value(values, api.key);
            computed.insert(*addr, value);
        }
    }

    // 2. Formeln evaluieren (in topologischer Reihenfolge)
    // Da wir keine Zyklen haben, können wir einfach alle Formeln durchgehen
    // Die Registry stellt sicher, dass Dependencies bereits berechnet sind
    let mut formula_addrs: Vec<CellAddr> = registry.formula_cells().iter().copied().collect();
    formula_addrs.sort(); // Sortiere nach Adresse für konsistente Reihenfolge

    for addr in formula_addrs {
        if let CellKind::Formula(f) = registry.get(addr) {
            let ctx = EvalContext {
                computed: &computed,
                api_values: values,
            };
            let result = (f.eval)(&ctx);
            computed.insert(addr, result);
        }
    }

    computed
}

/// Holt API-Wert aus ReportValues
///
/// Verwendet `get_owned()` um alle Keys inkl. Footer-Keys zu unterstützen.
fn get_api_value(values: &ReportValues, key: super::api::ApiKey) -> CellValue {
    values.get_owned(key)
}

/// Schreibt alle Zellen aus der Registry (API-Werte und Formeln mit Cache)
fn write_cells_from_registry(
    ws: &mut Worksheet,
    registry: &CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    computed: &HashMap<CellAddr, CellValue>,
    fmt: &FormatMatrix,
) -> Result<(), XlsxError> {
    // 1. API-Zellen schreiben (mit ihren berechneten Werten)
    for addr in registry.api_cells() {
        if let Some(value) = computed.get(addr) {
            if !value.is_empty() {
                write_cell_value(ws, *addr, value, fmt)?;
            }
        }
    }

    // 2. Formel-Zellen schreiben (mit gecachten Ergebnissen)
    for addr in registry.formula_cells() {
        if let CellKind::Formula(f) = registry.get(*addr) {
            let result = computed.get(addr).cloned().unwrap_or(CellValue::Empty);

            // Formula mit Result erstellen (Cache für Excel)
            let formula = Formula::new(f.excel).set_result(cell_value_to_string(&result));

            // Mit Format schreiben (locked)
            if let Some(format) = fmt.get_locked(addr.row, addr.col) {
                ws.write_formula_with_format(addr.row, addr.col, formula, &format)?;
            } else {
                let locked = Format::new().set_locked();
                ws.write_formula_with_format(addr.row, addr.col, formula, &locked)?;
            }
        }
    }

    Ok(())
}

/// Konvertiert CellValue zu String für Formula::set_result
fn cell_value_to_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(s) => s.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Date(d) => d.clone(),
    }
}

// ============================================================================
// Dynamic Section Support
// ============================================================================

/// Bereich für dynamische Zellen (kann wachsen)
#[derive(Debug, Clone)]
pub struct DynamicRange {
    /// Start-Zeile (0-basiert)
    pub start_row: u32,
    /// Aktuelle End-Zeile (wächst dynamisch)
    pub end_row: u32,
    /// Start-Spalte
    pub start_col: u16,
    /// End-Spalte
    pub end_col: u16,
}

impl DynamicRange {
    pub fn new(start_row: u32, start_col: u16, end_col: u16) -> Self {
        Self {
            start_row,
            end_row: start_row,
            start_col,
            end_col,
        }
    }

    /// Fügt eine Zeile hinzu
    pub fn add_row(&mut self) -> u32 {
        let row = self.end_row;
        self.end_row += 1;
        row
    }

    /// Anzahl der Zeilen
    pub fn row_count(&self) -> u32 {
        self.end_row - self.start_row
    }

    /// Alle Adressen im Bereich
    pub fn all_addrs(&self) -> Vec<CellAddr> {
        let mut addrs = Vec::new();
        for row in self.start_row..self.end_row {
            for col in self.start_col..=self.end_col {
                addrs.push(CellAddr::new(row, col));
            }
        }
        addrs
    }
}

/// Builder für dynamische Sections
pub struct DynamicSectionBuilder {
    /// Aktueller Bereich
    range: DynamicRange,
    /// Zeilen-Daten
    rows: Vec<DynamicRow>,
}

/// Eine dynamische Zeile
pub struct DynamicRow {
    /// Relative Zeilen-Nummer (0 = erste Zeile der Section)
    pub index: usize,
    /// Zell-Werte für diese Zeile
    pub cells: HashMap<u16, DynamicCellValue>,
}

/// Wert einer dynamischen Zelle
#[derive(Debug, Clone)]
pub enum DynamicCellValue {
    /// Statischer Wert
    Value(CellValue),
    /// Formel mit relativem Bezug
    Formula {
        /// Excel-Formel (mit Platzhaltern wie {row})
        template: String,
        /// Evaluator
        eval: DynamicEval,
    },
}

/// Evaluator für dynamische Formeln
#[derive(Debug, Clone)]
pub enum DynamicEval {
    /// Division: col_a / col_b
    Division { col_a: u16, col_b: u16 },
    /// Summe einer Spalte bis zur aktuellen Zeile
    RunningSum { col: u16 },
    /// Lookup in TEXT_MATRIX
    TextLookup { index: usize },
    /// Keine Auswertung (nur Excel-Formel)
    None,
}

impl DynamicSectionBuilder {
    pub fn new(start_row: u32, start_col: u16, end_col: u16) -> Self {
        Self {
            range: DynamicRange::new(start_row, start_col, end_col),
            rows: Vec::new(),
        }
    }

    /// Fügt eine Zeile hinzu
    pub fn add_row(&mut self) -> &mut DynamicRow {
        let index = self.rows.len();
        let _row = self.range.add_row();
        self.rows.push(DynamicRow {
            index,
            cells: HashMap::new(),
        });
        self.rows.last_mut().unwrap()
    }

    /// Gibt den aktuellen Bereich zurück
    pub fn range(&self) -> &DynamicRange {
        &self.range
    }

    /// Baut die dynamische Section
    pub fn build(self) -> DynamicSection {
        DynamicSection {
            range: self.range,
            rows: self.rows,
        }
    }
}

impl DynamicRow {
    /// Setzt einen Wert
    pub fn set_value(&mut self, col: u16, value: CellValue) -> &mut Self {
        self.cells.insert(col, DynamicCellValue::Value(value));
        self
    }

    /// Setzt eine Formel
    pub fn set_formula(&mut self, col: u16, template: &str, eval: DynamicEval) -> &mut Self {
        self.cells.insert(
            col,
            DynamicCellValue::Formula {
                template: template.to_string(),
                eval,
            },
        );
        self
    }
}

/// Eine gebaute dynamische Section
pub struct DynamicSection {
    pub range: DynamicRange,
    pub rows: Vec<DynamicRow>,
}

impl DynamicSection {
    /// Schreibt die Section ins Worksheet
    pub fn write(
        &self,
        ws: &mut Worksheet,
        fmt: &FormatMatrix,
        computed: &HashMap<CellAddr, CellValue>,
    ) -> Result<(), XlsxError> {
        for row_data in &self.rows {
            let abs_row = self.range.start_row + row_data.index as u32;

            for (&col, cell_value) in &row_data.cells {
                let addr = CellAddr::new(abs_row, col);

                match cell_value {
                    DynamicCellValue::Value(value) => {
                        write_cell_value(ws, addr, value, fmt)?;
                    }
                    DynamicCellValue::Formula { template, eval } => {
                        // Ersetze {row} Platzhalter
                        let excel_row = abs_row + 1;
                        let formula_str = template.replace("{row}", &excel_row.to_string());

                        // Evaluiere
                        let result = evaluate_dynamic(eval, abs_row, computed);

                        // Schreibe Formel
                        let formula =
                            Formula::new(&formula_str).set_result(cell_value_to_string(&result));

                        if let Some(format) = fmt.get_locked(abs_row, col) {
                            ws.write_formula_with_format(abs_row, col, formula, &format)?;
                        } else {
                            let locked = Format::new().set_locked();
                            ws.write_formula_with_format(abs_row, col, formula, &locked)?;
                        }
                    }
                }
            }
        }

        Ok(())
    }
}

/// Schreibt einen Zellwert
fn write_cell_value(
    ws: &mut Worksheet,
    addr: CellAddr,
    value: &CellValue,
    fmt: &FormatMatrix,
) -> Result<(), XlsxError> {
    let format = fmt.get(addr.row, addr.col);

    match value {
        CellValue::Empty => {
            if let Some(f) = format {
                ws.write_blank(addr.row, addr.col, &f)?;
            }
        }
        CellValue::Text(s) => {
            if let Some(f) = format {
                ws.write_string_with_format(addr.row, addr.col, s, &f)?;
            } else {
                ws.write_string(addr.row, addr.col, s)?;
            }
        }
        CellValue::Number(n) => {
            if let Some(f) = format {
                ws.write_number_with_format(addr.row, addr.col, *n, &f)?;
            } else {
                ws.write_number(addr.row, addr.col, *n)?;
            }
        }
        CellValue::Date(d) => {
            if let Some(f) = format {
                ws.write_string_with_format(addr.row, addr.col, d, &f)?;
            } else {
                ws.write_string(addr.row, addr.col, d)?;
            }
        }
    }

    Ok(())
}

/// Evaluiert eine dynamische Formel
fn evaluate_dynamic(
    eval: &DynamicEval,
    row: u32,
    computed: &HashMap<CellAddr, CellValue>,
) -> CellValue {
    match eval {
        DynamicEval::Division { col_a, col_b } => {
            let addr_a = CellAddr::new(row, *col_a);
            let addr_b = CellAddr::new(row, *col_b);

            let a = computed.get(&addr_a).and_then(|v| v.as_number());
            let b = computed.get(&addr_b).and_then(|v| v.as_number());

            match (a, b) {
                (Some(a), Some(b)) if b != 0.0 => CellValue::Number(a / b),
                _ => CellValue::Empty,
            }
        }
        DynamicEval::RunningSum { col } => {
            let mut sum = 0.0;
            for r in 0..=row {
                let addr = CellAddr::new(r, *col);
                if let Some(n) = computed.get(&addr).and_then(|v| v.as_number()) {
                    sum += n;
                }
            }
            CellValue::Number(sum)
        }
        DynamicEval::TextLookup { index } => {
            // Würde TEXT_MATRIX nutzen
            CellValue::Text(format!("Lookup {}", index))
        }
        DynamicEval::None => CellValue::Empty,
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dynamic_range() {
        let mut range = DynamicRange::new(31, 0, 21); // Start bei Row 32 (0-basiert: 31)

        assert_eq!(range.row_count(), 0);

        range.add_row();
        range.add_row();
        range.add_row();

        assert_eq!(range.row_count(), 3);
        assert_eq!(range.start_row, 31);
        assert_eq!(range.end_row, 34);
    }

    #[test]
    fn test_dynamic_section_builder() {
        let mut builder = DynamicSectionBuilder::new(31, 0, 21);

        {
            let row = builder.add_row();
            row.set_value(0, CellValue::Text("Test".to_string()));
            row.set_formula(1, "=A{row}*2", DynamicEval::None);
        }

        let section = builder.build();
        assert_eq!(section.rows.len(), 1);
        assert_eq!(section.range.row_count(), 1);
    }
}
