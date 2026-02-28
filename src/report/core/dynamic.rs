//! Dynamic Sections - Infrastruktur für variable Bereiche
//!
//! Der Finanzbericht hat:
//! - Statischer Teil: A1:V31 (Header, Tabelle, Panel)
//! - Dynamischer Teil: Ab Zeile 32, variable Anzahl Zeilen
//!
//! Dieses Modul stellt die Infrastruktur für den dynamischen Teil bereit.

use super::registry::CellAddr;
use crate::lang::data::TEXT_MATRIX;
use crate::report::api::{CellValue, ReportValues};
use std::collections::HashMap;

// ============================================================================
// Dynamic Registry Extension
// ============================================================================

/// Erweitert die Registry um dynamische Zellen
pub struct DynamicRegistry {
    /// Dynamische Formeln
    formulas: HashMap<CellAddr, DynamicFormula>,
    /// Dynamische API-Zellen
    api_cells: HashMap<CellAddr, DynamicApiKey>,
    /// Nächste freie Zeile
    next_row: u32,
}

/// Schlüssel für dynamische API-Zellen
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DynamicApiKey {
    /// Dynamische Zelle mit Zeilen-Index
    Cell { col: u16, row_index: u32 },
}

/// Dynamische Formel
pub struct DynamicFormula {
    /// Excel-Formel Template (mit {row} Platzhalter)
    pub template: &'static str,
    /// Input-Spalten (relativ zur gleichen Zeile)
    pub input_cols: Vec<u16>,
    /// Statische Werte
    pub statics: Vec<super::registry::StaticVal>,
    /// Sheet-Referenzen
    pub sheets: Vec<super::registry::SheetRef>,
    /// Abhängige Formeln in gleicher Zeile (Spalten)
    pub formula_dep_cols: Vec<u16>,
    /// Evaluator
    pub eval: Box<dyn Fn(&DynamicEvalContext) -> CellValue + Send + Sync>,
}

/// Kontext für dynamische Formel-Auswertung
pub struct DynamicEvalContext<'a> {
    /// Aktuelle Zeile
    pub row: u32,
    /// Alle berechneten Werte
    pub computed: &'a HashMap<CellAddr, CellValue>,
    /// API-Werte
    pub api_values: &'a ReportValues,
    /// Dynamische Werte
    pub dynamic_values: &'a HashMap<CellAddr, CellValue>,
}

impl<'a> DynamicEvalContext<'a> {
    /// Liest Wert aus gleicher Zeile
    pub fn cell_in_row(&self, col: u16) -> &CellValue {
        static EMPTY: CellValue = CellValue::Empty;
        let addr = CellAddr::new(self.row, col);
        self.dynamic_values
            .get(&addr)
            .or_else(|| self.computed.get(&addr))
            .unwrap_or(&EMPTY)
    }

    /// Liest Wert aus beliebiger Zelle
    pub fn cell(&self, addr: CellAddr) -> &CellValue {
        static EMPTY: CellValue = CellValue::Empty;
        self.dynamic_values
            .get(&addr)
            .or_else(|| self.computed.get(&addr))
            .unwrap_or(&EMPTY)
    }

    /// Sprache aus API
    pub fn language(&self) -> Option<&str> {
        self.api_values.language()
    }
}

impl DynamicRegistry {
    /// Erstellt eine neue dynamische Registry
    /// `start_row` ist die erste Zeile des dynamischen Bereichs (0-basiert)
    pub fn new(start_row: u32) -> Self {
        Self {
            formulas: HashMap::new(),
            api_cells: HashMap::new(),
            next_row: start_row,
        }
    }

    /// Reserviert die nächste Zeile
    pub fn next_row(&mut self) -> u32 {
        let row = self.next_row;
        self.next_row += 1;
        row
    }

    /// Aktuelle End-Zeile
    pub fn current_end_row(&self) -> u32 {
        self.next_row
    }

    /// Registriert eine dynamische API-Zelle
    pub fn register_api(&mut self, row: u32, col: u16) {
        let addr = CellAddr::new(row, col);
        self.api_cells.insert(
            addr,
            DynamicApiKey::Cell {
                col,
                row_index: row,
            },
        );
    }

    /// Registriert eine dynamische Formel
    pub fn register_formula(&mut self, row: u32, col: u16, formula: DynamicFormula) {
        let addr = CellAddr::new(row, col);
        self.formulas.insert(addr, formula);
    }

    /// Gibt alle dynamischen Adressen zurück
    pub fn all_addrs(&self) -> Vec<CellAddr> {
        let mut addrs: Vec<CellAddr> = self
            .api_cells
            .keys()
            .chain(self.formulas.keys())
            .copied()
            .collect();
        addrs.sort();
        addrs.dedup();
        addrs
    }

    /// Prüft ob Adresse dynamisch ist
    pub fn is_dynamic(&self, addr: CellAddr) -> bool {
        self.api_cells.contains_key(&addr) || self.formulas.contains_key(&addr)
    }
}

// ============================================================================
// Dynamic Section Definition
// ============================================================================

/// Definition einer dynamischen Section (z.B. Ausgaben-Tabelle)
pub struct DynamicSectionDef {
    /// Name der Section
    pub name: &'static str,
    /// Start-Spalte
    pub start_col: u16,
    /// End-Spalte
    pub end_col: u16,
    /// Spalten-Definitionen
    pub columns: Vec<DynamicColumnDef>,
}

/// Definition einer Spalte in einer dynamischen Section
pub struct DynamicColumnDef {
    /// Spalten-Index
    pub col: u16,
    /// Spalten-Typ
    pub kind: DynamicColumnKind,
}

/// Art einer dynamischen Spalte
pub enum DynamicColumnKind {
    /// API-Eingabe (Wert von außen)
    ApiInput,
    /// Text-Lookup
    TextLookup { index: usize },
    /// Formel mit Template
    Formula {
        template: &'static str,
        eval: Box<dyn Fn(&DynamicEvalContext) -> CellValue + Send + Sync>,
    },
    /// Statischer Wert pro Zeile (z.B. Zeilennummer)
    RowNumber,
    /// Summe aus anderen Spalten
    Sum { cols: Vec<u16> },
    /// Division
    Division {
        numerator_col: u16,
        denominator_col: u16,
    },
}

impl DynamicSectionDef {
    /// Erstellt eine neue Section-Definition
    pub fn new(name: &'static str, start_col: u16, end_col: u16) -> Self {
        Self {
            name,
            start_col,
            end_col,
            columns: Vec::new(),
        }
    }

    /// Fügt eine Spalte hinzu
    pub fn add_column(&mut self, col: u16, kind: DynamicColumnKind) -> &mut Self {
        self.columns.push(DynamicColumnDef { col, kind });
        self
    }

    /// Registriert eine Zeile dieser Section in der DynamicRegistry
    pub fn register_row(&self, registry: &mut DynamicRegistry, row: u32) {
        for col_def in &self.columns {
            match &col_def.kind {
                DynamicColumnKind::ApiInput => {
                    registry.register_api(row, col_def.col);
                }
                DynamicColumnKind::TextLookup { index } => {
                    let idx = *index;
                    registry.register_formula(
                        row,
                        col_def.col,
                        DynamicFormula {
                            template: "", // Wird dynamisch generiert
                            input_cols: vec![],
                            statics: vec![super::registry::StaticVal::Index(idx)],
                            sheets: vec![super::registry::SheetRef::lang_table()],
                            formula_dep_cols: vec![],
                            eval: Box::new(move |ctx| evaluate_text_lookup_dynamic(ctx, idx)),
                        },
                    );
                }
                DynamicColumnKind::Formula { template, eval: _ } => {
                    registry.register_formula(
                        row,
                        col_def.col,
                        DynamicFormula {
                            template,
                            input_cols: vec![],
                            statics: vec![],
                            sheets: vec![],
                            formula_dep_cols: vec![],
                            eval: Box::new(|_| CellValue::Empty),
                        },
                    );
                }
                DynamicColumnKind::RowNumber => {
                    // Zeilen-Nummer als Wert
                    registry.register_formula(
                        row,
                        col_def.col,
                        DynamicFormula {
                            template: "",
                            input_cols: vec![],
                            statics: vec![],
                            sheets: vec![],
                            formula_dep_cols: vec![],
                            eval: Box::new(move |ctx| CellValue::Number((ctx.row + 1) as f64)),
                        },
                    );
                }
                DynamicColumnKind::Sum { cols } => {
                    let cols_clone = cols.clone();
                    registry.register_formula(
                        row,
                        col_def.col,
                        DynamicFormula {
                            template: "",
                            input_cols: cols_clone.clone(),
                            statics: vec![],
                            sheets: vec![],
                            formula_dep_cols: vec![],
                            eval: Box::new(move |ctx| {
                                let mut sum = 0.0;
                                for &c in &cols_clone {
                                    if let Some(n) = ctx.cell_in_row(c).as_number() {
                                        sum += n;
                                    }
                                }
                                CellValue::Number(sum)
                            }),
                        },
                    );
                }
                DynamicColumnKind::Division {
                    numerator_col,
                    denominator_col,
                } => {
                    let num = *numerator_col;
                    let denom = *denominator_col;
                    registry.register_formula(
                        row,
                        col_def.col,
                        DynamicFormula {
                            template: "",
                            input_cols: vec![num, denom],
                            statics: vec![],
                            sheets: vec![],
                            formula_dep_cols: vec![],
                            eval: Box::new(move |ctx| {
                                let n = ctx.cell_in_row(num).as_number();
                                let d = ctx.cell_in_row(denom).as_number();
                                match (n, d) {
                                    (Some(n), Some(d)) if d != 0.0 => CellValue::Number(n / d),
                                    _ => CellValue::Empty,
                                }
                            }),
                        },
                    );
                }
            }
        }
    }
}

// ============================================================================
// Evaluation
// ============================================================================

/// Text-Lookup für dynamische Zellen
fn evaluate_text_lookup_dynamic(ctx: &DynamicEvalContext, index: usize) -> CellValue {
    let language = match ctx.language() {
        Some(lang) if !lang.is_empty() => lang,
        _ => return CellValue::Empty,
    };

    let lang_idx = TEXT_MATRIX
        .iter()
        .position(|row| !row.is_empty() && row[0].eq_ignore_ascii_case(language));

    let lang_idx = match lang_idx {
        Some(idx) => idx,
        None => return CellValue::Empty,
    };

    let text_idx = index.saturating_sub(1);

    TEXT_MATRIX
        .get(lang_idx)
        .and_then(|row| row.get(text_idx))
        .map(|s| CellValue::Text(s.to_string()))
        .unwrap_or(CellValue::Empty)
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dynamic_registry() {
        let mut registry = DynamicRegistry::new(31); // Start nach Zeile 31 (V31)

        let row = registry.next_row();
        assert_eq!(row, 31);

        registry.register_api(row, 0); // A32
        registry.register_api(row, 1); // B32

        assert!(registry.is_dynamic(CellAddr::new(31, 0)));
        assert!(registry.is_dynamic(CellAddr::new(31, 1)));
        assert!(!registry.is_dynamic(CellAddr::new(30, 0))); // V31 ist nicht dynamisch
    }

    #[test]
    fn test_dynamic_section_def() {
        let mut section = DynamicSectionDef::new("Ausgaben", 0, 7);
        section
            .add_column(0, DynamicColumnKind::RowNumber)
            .add_column(1, DynamicColumnKind::TextLookup { index: 10 })
            .add_column(2, DynamicColumnKind::ApiInput)
            .add_column(3, DynamicColumnKind::ApiInput)
            .add_column(4, DynamicColumnKind::Sum { cols: vec![2, 3] });

        assert_eq!(section.columns.len(), 5);
    }
}
