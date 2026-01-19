//! Formel-Definitionen und Auswertung für den Finanzbericht
//!
//! Jede Formel deklariert:
//! - Ihre Zelladresse
//! - Die Abhängigkeiten (welche Zellen sie benötigt)
//! - Die Excel-Formel (für die Datei)
//! - Die Rust-Evaluierungsfunktion (für Berechnungen)

use super::cells::CellAddress;
use super::values::{CellValue, ReportValues};
use crate::v2::Sprachversion::data::TEXT_MATRIX;
use once_cell::sync::Lazy;
use std::collections::HashMap;

// ============================================================================
// Formula Definition
// ============================================================================

/// Definition einer Formel
#[derive(Debug, Clone)]
pub struct FormulaDefinition {
    /// Zelladresse wo die Formel steht
    pub address: CellAddress,
    /// Excel-Formel als String
    pub excel_formula: String,
    /// Abhängigkeiten: Welche Zellen werden für die Berechnung benötigt
    pub dependencies: Vec<CellAddress>,
    /// Formel-Typ für die Auswertung
    pub formula_type: FormulaType,
}

/// Typ der Formel für die Auswertung
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaType {
    /// VLOOKUP in TEXT_MATRIX: =IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE))
    TextLookup { index: usize },
    /// VLOOKUP mit Default-Text wenn leer
    TextLookupWithDefault { index: usize, default: String },
    /// Hyperlink zu VLOOKUP
    HyperlinkLookup { index: usize },
    /// Währung aus E3 oder VLOOKUP wenn leer
    CurrencyOrLookup { index: usize },
    /// Einfache Zellreferenz (kopiert Wert)
    CellReference { source: CellAddress },
    /// IFERROR Division
    DivisionWithError {
        numerator: CellAddress,
        denominator: CellAddress,
    },
    /// SUMPRODUCT mit ROUND
    SumProductRound {
        range_start: CellAddress,
        range_end: CellAddress,
    },
    /// Komplexe Formel (manuell implementiert)
    Custom,
    /// Right Panel Calc: =IF(M{row}="","",N{row}/M{row})
    RightPanelCalc { amount1_col: u16, amount2_col: u16 },
}

// ============================================================================
// Formula Builders - Vereinfachte Erstellung von Formeln
// ============================================================================

/// E2 Adresse (Sprachauswahl)
const E2: CellAddress = CellAddress::new(1, 4);
/// E3 Adresse (Währung)
const E3: CellAddress = CellAddress::new(2, 4);

/// Erstellt eine TextLookup Formel
fn text_lookup(row: u32, col: u16, index: usize) -> FormulaDefinition {
    FormulaDefinition {
        address: CellAddress::new(row, col),
        excel_formula: format!(
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
            index
        ),
        dependencies: vec![E2],
        formula_type: FormulaType::TextLookup { index },
    }
}

/// Erstellt eine TextLookup Formel mit Default
fn text_lookup_default(row: u32, col: u16, index: usize, default: &str) -> FormulaDefinition {
    FormulaDefinition {
        address: CellAddress::new(row, col),
        excel_formula: format!(
            r#"=IF($E$2="","{}",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
            default, index
        ),
        dependencies: vec![E2],
        formula_type: FormulaType::TextLookupWithDefault {
            index,
            default: default.to_string(),
        },
    }
}

/// Erstellt eine Hyperlink-Lookup Formel
fn hyperlink_lookup(row: u32, col: u16, index: usize) -> FormulaDefinition {
    FormulaDefinition {
        address: CellAddress::new(row, col),
        excel_formula: format!(
            r#"=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
            index
        ),
        dependencies: vec![E2],
        formula_type: FormulaType::HyperlinkLookup { index },
    }
}

/// Erstellt eine Währung-oder-Lookup Formel
fn currency_or_lookup(row: u32, col: u16, index: usize) -> FormulaDefinition {
    FormulaDefinition {
        address: CellAddress::new(row, col),
        excel_formula: format!(
            r#"=IF(E3="",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE),E3)"#,
            index
        ),
        dependencies: vec![E2, E3],
        formula_type: FormulaType::CurrencyOrLookup { index },
    }
}

/// Erstellt eine IFERROR Division Formel
fn division_error(
    row: u32,
    col: u16,
    num_row: u32,
    num_col: u16,
    den_row: u32,
    den_col: u16,
) -> FormulaDefinition {
    let num_addr = CellAddress::new(num_row, num_col);
    let den_addr = CellAddress::new(den_row, den_col);
    FormulaDefinition {
        address: CellAddress::new(row, col),
        excel_formula: format!(
            "=IFERROR({}/{},0)",
            num_addr.to_excel_notation(),
            den_addr.to_excel_notation()
        ),
        dependencies: vec![num_addr, den_addr],
        formula_type: FormulaType::DivisionWithError {
            numerator: num_addr,
            denominator: den_addr,
        },
    }
}

/// Erstellt eine SUMPRODUCT ROUND Formel
fn sum_product_round(
    row: u32,
    col: u16,
    start_row: u32,
    end_row: u32,
    data_col: u16,
) -> FormulaDefinition {
    let start = CellAddress::new(start_row, data_col);
    let end = CellAddress::new(end_row, data_col);
    FormulaDefinition {
        address: CellAddress::new(row, col),
        excel_formula: format!(
            "=SUMPRODUCT(ROUND({}:{}, 2))",
            start.to_excel_notation(),
            end.to_excel_notation()
        ),
        dependencies: vec![], // Dynamische Abhängigkeiten
        formula_type: FormulaType::SumProductRound {
            range_start: start,
            range_end: end,
        },
    }
}

// ============================================================================
// Formula Cache
// ============================================================================

/// Cache für berechnete Formelwerte
#[derive(Debug, Default)]
pub struct FormulaCache {
    cache: HashMap<CellAddress, CellValue>,
}

impl FormulaCache {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn get(&self, address: &CellAddress) -> Option<&CellValue> {
        self.cache.get(address)
    }

    pub fn set(&mut self, address: CellAddress, value: CellValue) {
        self.cache.insert(address, value);
    }

    pub fn clear(&mut self) {
        self.cache.clear();
    }
}

// ============================================================================
// Formula Evaluator
// ============================================================================

/// Evaluiert eine Formel basierend auf den Eingabewerten
pub fn evaluate_formula(
    formula: &FormulaDefinition,
    values: &ReportValues,
    cache: &mut FormulaCache,
) -> CellValue {
    // Prüfe Cache
    if let Some(cached) = cache.get(&formula.address) {
        return cached.clone();
    }

    let result = match &formula.formula_type {
        FormulaType::TextLookup { index } => evaluate_text_lookup(values, *index),
        FormulaType::TextLookupWithDefault { index, default } => {
            evaluate_text_lookup_with_default(values, *index, default)
        }
        FormulaType::HyperlinkLookup { index } => evaluate_hyperlink_lookup(values, *index),
        FormulaType::CurrencyOrLookup { index } => evaluate_currency_or_lookup(values, *index),
        FormulaType::CellReference { source } => {
            cache.get(source).cloned().unwrap_or(CellValue::Empty)
        }
        FormulaType::DivisionWithError {
            numerator,
            denominator,
        } => evaluate_division(cache, numerator, denominator),
        FormulaType::SumProductRound {
            range_start,
            range_end,
        } => evaluate_sum_product(cache, range_start, range_end),
        FormulaType::RightPanelCalc {
            amount1_col,
            amount2_col,
        } => {
            let row = formula.address.row;
            let amount1_addr = CellAddress::new(row, *amount1_col);
            let amount2_addr = CellAddress::new(row, *amount2_col);
            evaluate_division(cache, &amount2_addr, &amount1_addr)
        }
        FormulaType::Custom => CellValue::Empty,
    };

    // Cache das Ergebnis
    cache.set(formula.address, result.clone());
    result
}

// ============================================================================
// Lookup Functions
// ============================================================================

/// Findet den Sprach-Index in TEXT_MATRIX basierend auf dem Sprachcode
fn find_language_index(language: &str) -> Option<usize> {
    TEXT_MATRIX
        .iter()
        .position(|row| !row.is_empty() && row[0].eq_ignore_ascii_case(language))
}

/// Evaluiert: =IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE))
fn evaluate_text_lookup(values: &ReportValues, index: usize) -> CellValue {
    let language = match values.language() {
        Some(lang) if !lang.is_empty() => lang,
        _ => return CellValue::Empty,
    };

    let lang_idx = match find_language_index(language) {
        Some(idx) => idx,
        None => return CellValue::Empty,
    };

    // Index ist 1-basiert in Excel, aber 0-basiert in TEXT_MATRIX
    // VLOOKUP index 2 = TEXT_MATRIX[lang][1]
    let text_idx = index.saturating_sub(1);

    TEXT_MATRIX
        .get(lang_idx)
        .and_then(|row| row.get(text_idx))
        .map(|s| CellValue::Text(s.to_string()))
        .unwrap_or(CellValue::Empty)
}

/// Evaluiert: =IF($E$2="",default,VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE))
fn evaluate_text_lookup_with_default(
    values: &ReportValues,
    index: usize,
    default: &str,
) -> CellValue {
    let language = match values.language() {
        Some(lang) if !lang.is_empty() => lang,
        _ => return CellValue::Text(default.to_string()),
    };

    let lang_idx = match find_language_index(language) {
        Some(idx) => idx,
        None => return CellValue::Text(default.to_string()),
    };

    let text_idx = index.saturating_sub(1);

    TEXT_MATRIX
        .get(lang_idx)
        .and_then(|row| row.get(text_idx))
        .map(|s| CellValue::Text(s.to_string()))
        .unwrap_or(CellValue::Text(default.to_string()))
}

/// Evaluiert: =HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE))
fn evaluate_hyperlink_lookup(values: &ReportValues, index: usize) -> CellValue {
    evaluate_text_lookup(values, index)
}

/// Evaluiert: =IF(E3="",VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE),E3)
fn evaluate_currency_or_lookup(values: &ReportValues, index: usize) -> CellValue {
    if let Some(currency) = values.currency() {
        if !currency.is_empty() {
            return CellValue::Text(currency.to_string());
        }
    }
    evaluate_text_lookup(values, index)
}

/// Evaluiert Division mit IFERROR: =IFERROR(A/B, 0)
fn evaluate_division(
    cache: &FormulaCache,
    numerator: &CellAddress,
    denominator: &CellAddress,
) -> CellValue {
    let num = cache
        .get(numerator)
        .and_then(|v| v.as_number())
        .unwrap_or(0.0);

    let denom = cache
        .get(denominator)
        .and_then(|v| v.as_number())
        .unwrap_or(0.0);

    if denom == 0.0 {
        CellValue::Number(0.0)
    } else {
        CellValue::Number(num / denom)
    }
}

/// Evaluiert SUMPRODUCT mit ROUND
fn evaluate_sum_product(
    cache: &FormulaCache,
    range_start: &CellAddress,
    range_end: &CellAddress,
) -> CellValue {
    let mut sum = 0.0;

    for row in range_start.row..=range_end.row {
        let addr = CellAddress::new(row, range_start.col);
        if let Some(value) = cache.get(&addr) {
            if let Some(num) = value.as_number() {
                sum += (num * 100.0).round() / 100.0;
            }
        }
    }

    CellValue::Number(sum)
}

// ============================================================================
// All Header Formulas - Zentrale Definition
// ============================================================================

/// Alle Formeln im Header-Bereich
pub static HEADER_FORMULAS: Lazy<Vec<FormulaDefinition>> = Lazy::new(|| {
    vec![
        // Row 0 (B1, J1)
        text_lookup(0, 1, 2),  // B1: Titel
        text_lookup(0, 9, 60), // J1: Ausfüllbare Felder Info
        // Row 1 (D2, J2)
        text_lookup_default(1, 3, 27, "Chose your language"), // D2: Sprache Label
        text_lookup(1, 9, 61),                                // J2: Tipp
        // Row 2 (B3, D3)
        text_lookup(2, 1, 3),  // B3: Übersicht Label
        text_lookup(2, 3, 28), // D3: Lokalwährung Label
        // Row 3 (J4 - Hyperlink)
        hyperlink_lookup(3, 9, 62), // J4: Website Link
        // Row 4 (B5)
        text_lookup(4, 1, 4), // B5: Projektnummer Label
        // Row 5 (B6)
        text_lookup(5, 1, 26), // B6: Projekttitel Label
        // Row 7 (B8, D8, F8)
        text_lookup(7, 1, 5), // B8: Projektlaufzeit Label
        text_lookup(7, 3, 7), // D8: von Label
        text_lookup(7, 5, 8), // F8: bis Label
        // Row 8 (B9, D9, F9)
        text_lookup(8, 1, 6), // B9: Berichtszeitraum Label
        text_lookup(8, 3, 7), // D9: von Label
        text_lookup(8, 5, 8), // F9: bis Label
        // K8 (Wechselkurs-Label)
        text_lookup(7, 10, 59), // K8: WK Label
        // Row 10 (J11, Q11, D11-H11)
        text_lookup(10, 9, 19),  // J11: KMW Mittel
        text_lookup(10, 16, 19), // Q11: KMW Mittel
        text_lookup(10, 3, 11),  // D11: Bewilligtes Budget
        text_lookup(10, 4, 12),  // E11: Einnahmen Berichtszeitraum
        text_lookup(10, 5, 13),  // F11: Einnahmen gesamt
        text_lookup(10, 6, 14),  // G11: % erhaltene Einnahmen
        text_lookup(10, 7, 15),  // H11: Begründung
        // Row 11 (B12)
        text_lookup(11, 1, 9), // B12: EINNAHMEN
        // Row 12 (B13, L13, N13, O13, S13, U13, V13)
        text_lookup(12, 1, 10),         // B13: (in lokaler Währung)
        text_lookup(12, 11, 22),        // L13: Datum
        currency_or_lookup(12, 13, 28), // N13: Währung
        text_lookup(12, 14, 58),        // O13: Wechselkurs
        text_lookup(12, 18, 22),        // S13: Datum
        currency_or_lookup(12, 20, 28), // U13: Währung
        text_lookup(12, 21, 58),        // V13: Wechselkurs
        // Table Body Labels (B15-B19)
        text_lookup(14, 1, 16), // B15: Saldovortrag
        text_lookup(15, 1, 17), // B16: Lokale Eigenleistung
        text_lookup(16, 1, 18), // B17: Beiträge dritter Seite
        text_lookup(17, 1, 19), // B18: KMW Mittel
        text_lookup(18, 1, 20), // B19: Zinserträge
        // Table Body % Formulas (G16-G19)
        division_error(15, 6, 15, 5, 15, 3), // G16: F16/D16
        division_error(16, 6, 16, 5, 16, 3), // G17: F17/D17
        division_error(17, 6, 17, 5, 17, 3), // G18: F18/D18
        division_error(18, 6, 18, 5, 18, 3), // G19: F19/D19
        // Summary Row (B20, D20-G20)
        text_lookup(19, 1, 21),              // B20: GESAMT
        sum_product_round(19, 3, 14, 18, 3), // D20: SUM(D15:D19)
        sum_product_round(19, 4, 14, 18, 4), // E20: SUM(E15:E19)
        sum_product_round(19, 5, 14, 18, 5), // F20: SUM(F15:F19)
        division_error(19, 6, 19, 5, 19, 3), // G20: F20/D20
    ]
});

/// Generiert Formeln für das Right Panel (K14-K31 und R14-R31 Text, O14-O31 und V14-V31 Calc)
pub fn generate_right_panel_formulas() -> Vec<FormulaDefinition> {
    let mut formulas = Vec::new();

    for i in 0..18 {
        let row = 13 + i as u32;
        let excel_row = row + 1;

        // K (Col 10): Text Formel
        formulas.push(text_lookup(row, 10, 23));

        // O (Col 14): Calc Formel
        formulas.push(FormulaDefinition {
            address: CellAddress::new(row, 14),
            excel_formula: format!(
                "=IF(M{}=\"\",\"\",N{}/M{})",
                excel_row, excel_row, excel_row
            ),
            dependencies: vec![
                CellAddress::new(row, 12), // M
                CellAddress::new(row, 13), // N
            ],
            formula_type: FormulaType::RightPanelCalc {
                amount1_col: 12,
                amount2_col: 13,
            },
        });

        // R (Col 17): Text Formel
        formulas.push(text_lookup(row, 17, 23));

        // V (Col 21): Calc Formel
        formulas.push(FormulaDefinition {
            address: CellAddress::new(row, 21),
            excel_formula: format!(
                "=IF(T{}=\"\",\"\",U{}/T{})",
                excel_row, excel_row, excel_row
            ),
            dependencies: vec![
                CellAddress::new(row, 19), // T
                CellAddress::new(row, 20), // U
            ],
            formula_type: FormulaType::RightPanelCalc {
                amount1_col: 19,
                amount2_col: 20,
            },
        });
    }

    formulas
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_lookup_german() {
        let values = ReportValues::new().with_language("deutsch");
        let result = evaluate_text_lookup(&values, 2);
        assert_eq!(result.as_text(), Some("F I N A N Z B E R I C H T"));
    }

    #[test]
    fn test_text_lookup_english() {
        let values = ReportValues::new().with_language("english");
        let result = evaluate_text_lookup(&values, 2);
        assert_eq!(result.as_text(), Some("F I N A N C I A L  R E P O R T "));
    }

    #[test]
    fn test_text_lookup_empty_language() {
        let values = ReportValues::new();
        let result = evaluate_text_lookup(&values, 2);
        assert!(result.is_empty());
    }

    #[test]
    fn test_currency_or_lookup_with_currency() {
        let values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("USD");
        let result = evaluate_currency_or_lookup(&values, 28);
        assert_eq!(result.as_text(), Some("USD"));
    }

    #[test]
    fn test_currency_or_lookup_without_currency() {
        let values = ReportValues::new().with_language("deutsch");
        let result = evaluate_currency_or_lookup(&values, 28);
        assert!(result.as_text().is_some());
    }

    #[test]
    fn test_header_formulas_count() {
        // Stelle sicher, dass alle Formeln generiert wurden
        assert!(!HEADER_FORMULAS.is_empty());
    }

    #[test]
    fn test_right_panel_formulas() {
        let formulas = generate_right_panel_formulas();
        // 18 Zeilen * 4 Formeln pro Zeile = 72 Formeln
        assert_eq!(formulas.len(), 72);
    }
}
