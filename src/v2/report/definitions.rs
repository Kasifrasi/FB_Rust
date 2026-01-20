//! Zell-Definitionen für den Finanzbericht
//!
//! Hier werden ALLE Zellen im Bereich A1:V31 registriert.
//! Jede Zelle gehört zu genau einer Kategorie:
//! - Api: Wert kommt von außen (ReportValues) - definiert in api.rs
//! - Static: Fester Wert im Code (aktuell keine)
//! - Formula: Hat Excel-Formel + Rust-Evaluator
//! - Empty: Nicht registriert = Standardwert

use super::api::register_all_api_cells;
use super::registry::{
    addr, builder::FormulaBuilder, CellAddr, CellRegistry, EvalContext, FormulaCell,
    FormulaCellDeps, Inputs, RegistryError, Sheets, Statics,
};
use super::values::CellValue;
use crate::v2::lang::data::TEXT_MATRIX;

// ============================================================================
// Öffentliche VLOOKUP-Evaluierung
// ============================================================================

/// Evaluiert einen VLOOKUP-Index für eine gegebene Sprache
///
/// Dies ist die öffentliche Version von `evaluate_text_lookup`, die direkt
/// mit der Sprache als Parameter arbeitet (ohne EvalContext).
///
/// # Arguments
/// * `language` - Die Sprache (z.B. "deutsch", "english")
/// * `index` - Der VLOOKUP-Index (1-basiert, wie in Excel)
///
/// # Returns
/// Der Text aus TEXT_MATRIX oder Empty wenn nicht gefunden
pub fn lookup_text(language: Option<&str>, index: usize) -> CellValue {
    let language = match language {
        Some(lang) if !lang.is_empty() => lang,
        _ => return CellValue::Empty,
    };

    let lang_idx = match find_language_index(language) {
        Some(idx) => idx,
        None => return CellValue::Empty,
    };

    // Index ist 1-basiert in Excel, aber 0-basiert in TEXT_MATRIX
    let text_idx = index.saturating_sub(1);

    TEXT_MATRIX
        .get(lang_idx)
        .and_then(|row| row.get(text_idx))
        .map(|s| CellValue::Text(s.to_string()))
        .unwrap_or(CellValue::Empty)
}

/// Evaluiert einen VLOOKUP-Index und gibt den Text als String zurück
///
/// Convenience-Funktion für Fälle wo nur der String-Wert benötigt wird.
pub fn lookup_text_string(language: Option<&str>, index: usize) -> Option<String> {
    match lookup_text(language, index) {
        CellValue::Text(s) => Some(s),
        _ => None,
    }
}

// ============================================================================
// Registry Builder
// ============================================================================

/// Erstellt eine vollständig konfigurierte Registry
pub fn build_registry(
) -> Result<CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>, RegistryError> {
    let mut registry = CellRegistry::new();

    // 1. API-Zellen registrieren (aus api.rs - EINZIGE QUELLE DER WAHRHEIT)
    register_all_api_cells(&mut registry)?;

    // 2. Statische Zellen registrieren (aktuell keine)
    // register_static_cells(&mut registry)?;

    // 3. Formel-Zellen registrieren (IN ABHÄNGIGKEITSREIHENFOLGE!)
    register_formula_cells(&mut registry)?;

    Ok(registry)
}

// ============================================================================
// Formula Cells
// ============================================================================

fn register_formula_cells(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
) -> Result<(), RegistryError> {
    // ========================================================================
    // Row 0: B1 (Titel), J1 (Ausfüllbare Felder Info)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(0, 1), 2)?; // B1
    register_text_lookup(registry, CellAddr::new(0, 9), 60)?; // J1

    // ========================================================================
    // Row 1: D2 (Sprache Label), J2 (Tipp)
    // ========================================================================

    register_text_lookup_default(registry, CellAddr::new(1, 3), 27, "Chose your language")?; // D2
    register_text_lookup(registry, CellAddr::new(1, 9), 61)?; // J2

    // ========================================================================
    // Row 2: B3 (Übersicht), D3 (Lokalwährung Label)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(2, 1), 3)?; // B3
    register_text_lookup(registry, CellAddr::new(2, 3), 28)?; // D3

    // ========================================================================
    // Row 3: J4 (Hyperlink)
    // ========================================================================

    register_hyperlink_lookup(registry, CellAddr::new(3, 9), 62)?; // J4

    // ========================================================================
    // Row 4: B5 (Projektnummer Label)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(4, 1), 4)?; // B5

    // ========================================================================
    // Row 5: B6 (Projekttitel Label)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(5, 1), 26)?; // B6

    // ========================================================================
    // Row 7: B8 (Projektlaufzeit), D8 (von), F8 (bis)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(7, 1), 5)?; // B8
    register_text_lookup(registry, CellAddr::new(7, 3), 7)?; // D8
    register_text_lookup(registry, CellAddr::new(7, 5), 8)?; // F8

    // ========================================================================
    // Row 8: B9 (Berichtszeitraum), D9 (von), F9 (bis), K8 (WK Label)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(8, 1), 6)?; // B9
    register_text_lookup(registry, CellAddr::new(8, 3), 7)?; // D9
    register_text_lookup(registry, CellAddr::new(8, 5), 8)?; // F9
    register_text_lookup(registry, CellAddr::new(7, 10), 59)?; // K8

    // ========================================================================
    // Row 10: J11 (KMW Mittel), Q11, D11-H11 (Tabellenkopf)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(10, 9), 19)?; // J11
    register_text_lookup(registry, CellAddr::new(10, 16), 19)?; // Q11
    register_text_lookup(registry, CellAddr::new(10, 3), 11)?; // D11
    register_text_lookup(registry, CellAddr::new(10, 4), 12)?; // E11
    register_text_lookup(registry, CellAddr::new(10, 5), 13)?; // F11
    register_text_lookup(registry, CellAddr::new(10, 6), 14)?; // G11
    register_text_lookup(registry, CellAddr::new(10, 7), 15)?; // H11

    // ========================================================================
    // Row 11: B12 (EINNAHMEN)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(11, 1), 9)?; // B12

    // ========================================================================
    // Row 12: B13, L13, M13, N13, O13, S13, T13, U13, V13
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(12, 1), 10)?; // B13

    // Right Panel Header - Left Side
    register_text_lookup(registry, CellAddr::new(12, 11), 22)?; // L13: Datum
    register_text_lookup(registry, CellAddr::new(12, 12), 63)?; // M13: Euro
    register_currency_or_lookup(registry, CellAddr::new(12, 13), 28)?; // N13: Währung
    register_text_lookup(registry, CellAddr::new(12, 14), 58)?; // O13: Wechselkurs

    // Right Panel Header - Right Side
    register_text_lookup(registry, CellAddr::new(12, 18), 22)?; // S13: Datum
    register_text_lookup(registry, CellAddr::new(12, 19), 63)?; // T13: Euro
    register_currency_or_lookup(registry, CellAddr::new(12, 20), 28)?; // U13: Währung
    register_text_lookup(registry, CellAddr::new(12, 21), 58)?; // V13: Wechselkurs

    // ========================================================================
    // Row 14-18: B15-B19 (Tabellen-Labels)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(14, 1), 16)?; // B15: Saldovortrag
    register_text_lookup(registry, CellAddr::new(15, 1), 17)?; // B16: Lokale Eigenleistung
    register_text_lookup(registry, CellAddr::new(16, 1), 18)?; // B17: Beiträge dritter Seite
    register_text_lookup(registry, CellAddr::new(17, 1), 19)?; // B18: KMW Mittel
    register_text_lookup(registry, CellAddr::new(18, 1), 20)?; // B19: Zinserträge

    // ========================================================================
    // Row 15-19: G16-G19 (Prozent-Formeln) - IFERROR(F/D,0)
    // HINWEIS: G15 ist leer oder auch Formel - hier nur G16-G19
    // ========================================================================

    for i in 1..5u8 {
        // G16-G19 (rows 15-18, col 6)
        register_iferror_division(
            registry,
            CellAddr::new(14 + i as u32, 6), // G16-G19
            CellAddr::new(14 + i as u32, 5), // F16-F19 (numerator)
            CellAddr::new(14 + i as u32, 3), // D16-D19 (denominator)
        )?;
    }

    // ========================================================================
    // Row 19: B20 (GESAMT), D20-G20 (Summen und Prozent)
    // ========================================================================

    register_text_lookup(registry, CellAddr::new(19, 1), 21)?; // B20

    // D20: =SUMPRODUCT(ROUND(D15:D19,2))
    register_sumproduct_round(
        registry,
        CellAddr::new(19, 3), // D20
        CellAddr::new(14, 3), // D15
        CellAddr::new(18, 3), // D19
    )?;

    // E20: =SUMPRODUCT(ROUND(E15:E19,2))
    register_sumproduct_round(
        registry,
        CellAddr::new(19, 4), // E20
        CellAddr::new(14, 4), // E15
        CellAddr::new(18, 4), // E19
    )?;

    // F20: =SUMPRODUCT(ROUND(F15:F19,2))
    register_sumproduct_round(
        registry,
        CellAddr::new(19, 5), // F20
        CellAddr::new(14, 5), // F15
        CellAddr::new(18, 5), // F19
    )?;

    // G20: =IFERROR(F20/D20,0) - hängt von D20 und F20 ab (beides Formeln!)
    register_iferror_division_formula_deps(
        registry,
        CellAddr::new(19, 6), // G20
        CellAddr::new(19, 5), // F20 (numerator) - ist Formel!
        CellAddr::new(19, 3), // D20 (denominator) - ist Formel!
    )?;

    // ========================================================================
    // Right Panel: O14-O31, V14-V31 (Berechnete Wechselkurs-Spalten)
    // ========================================================================

    for i in 0..18u8 {
        let row = 13 + i as u32;

        // O-Spalte (Col 14): =IF(M{row}="","",N{row}/M{row})
        register_right_panel_calc(registry, CellAddr::new(row, 14), 12, 13)?;

        // V-Spalte (Col 21): =IF(T{row}="","",U{row}/T{row})
        register_right_panel_calc(registry, CellAddr::new(row, 21), 19, 20)?;
    }

    Ok(())
}

// ============================================================================
// Formula Registration Helpers
// ============================================================================

/// Registriert eine TextLookup Formel: =IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE))
fn register_text_lookup(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    index: usize,
) -> Result<(), RegistryError> {
    let excel = Box::leak(
        format!(
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
            index
        )
        .into_boxed_str(),
    );

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::one(addr::E2))
        .statics(Statics::index(index))
        .sheets(Sheets::lang_table())
        .formula_deps(FormulaCellDeps::none())
        .eval(move |ctx| evaluate_text_lookup(ctx, index))
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

/// Registriert eine TextLookup Formel mit Default
fn register_text_lookup_default(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    index: usize,
    default: &'static str,
) -> Result<(), RegistryError> {
    let excel = Box::leak(
        format!(
            r#"=IF($E$2="","{}",VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
            default, index
        )
        .into_boxed_str(),
    );

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::one(addr::E2))
        .statics(Statics::index(index))
        .sheets(Sheets::lang_table())
        .formula_deps(FormulaCellDeps::none())
        .eval(move |ctx| evaluate_text_lookup_default(ctx, index, default))
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

/// Registriert eine Hyperlink Formel
fn register_hyperlink_lookup(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    index: usize,
) -> Result<(), RegistryError> {
    let excel = Box::leak(
        format!(
            r#"=HYPERLINK(VLOOKUP($E$2,Sprachversionen!$B:$BN,{},FALSE))"#,
            index
        )
        .into_boxed_str(),
    );

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::one(addr::E2))
        .statics(Statics::index(index))
        .sheets(Sheets::lang_table())
        .formula_deps(FormulaCellDeps::none())
        .eval(move |ctx| evaluate_text_lookup(ctx, index))
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

/// Registriert eine Currency-or-Lookup Formel: =IF(E3="",VLOOKUP(...),E3)
fn register_currency_or_lookup(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    index: usize,
) -> Result<(), RegistryError> {
    let excel: &'static str = r#"=IF(E3="",VLOOKUP($E$2,Sprachversionen!$B:$BN,28,FALSE),E3)"#;

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::many(vec![addr::E2, addr::E3]))
        .statics(Statics::index(index))
        .sheets(Sheets::lang_table())
        .formula_deps(FormulaCellDeps::none())
        .eval(move |ctx| evaluate_currency_or_lookup(ctx, index))
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

/// Registriert eine IFERROR Division Formel: =IFERROR(numerator/denominator,0)
/// Verwendet wenn numerator und denominator API-Zellen sind
fn register_iferror_division(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    numerator: CellAddr,
    denominator: CellAddr,
) -> Result<(), RegistryError> {
    let excel = Box::leak(
        format!(
            "=IFERROR({}/{},0)",
            numerator.to_excel(),
            denominator.to_excel()
        )
        .into_boxed_str(),
    );

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::many(vec![numerator, denominator]))
        .statics(Statics::none())
        .sheets(Sheets::none())
        .formula_deps(FormulaCellDeps::none())
        .eval(move |ctx| {
            let num = ctx.cell(numerator).as_number();
            let denom = ctx.cell(denominator).as_number();

            match (num, denom) {
                (Some(n), Some(d)) if d != 0.0 => CellValue::Number(n / d),
                _ => CellValue::Number(0.0), // IFERROR returns 0
            }
        })
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

/// Registriert eine IFERROR Division Formel wo numerator/denominator selbst Formeln sind
fn register_iferror_division_formula_deps(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    numerator: CellAddr,
    denominator: CellAddr,
) -> Result<(), RegistryError> {
    let excel = Box::leak(
        format!(
            "=IFERROR({}/{},0)",
            numerator.to_excel(),
            denominator.to_excel()
        )
        .into_boxed_str(),
    );

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::none()) // Keine direkten API-Inputs
        .statics(Statics::none())
        .sheets(Sheets::none())
        .formula_deps(FormulaCellDeps::many(vec![numerator, denominator])) // Abhängig von anderen Formeln!
        .eval(move |ctx| {
            let num = ctx.cell(numerator).as_number();
            let denom = ctx.cell(denominator).as_number();

            match (num, denom) {
                (Some(n), Some(d)) if d != 0.0 => CellValue::Number(n / d),
                _ => CellValue::Number(0.0), // IFERROR returns 0
            }
        })
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

/// Registriert eine SUMPRODUCT(ROUND(...,2)) Formel
fn register_sumproduct_round(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    range_start: CellAddr,
    range_end: CellAddr,
) -> Result<(), RegistryError> {
    let excel = Box::leak(
        format!(
            "=SUMPRODUCT(ROUND({}:{},2))",
            range_start.to_excel(),
            range_end.to_excel()
        )
        .into_boxed_str(),
    );

    // Sammle alle Zellen im Bereich als Inputs
    let mut inputs = Vec::new();
    for row in range_start.row..=range_end.row {
        inputs.push(CellAddr::new(row, range_start.col));
    }

    let start_row = range_start.row;
    let end_row = range_end.row;
    let col = range_start.col;

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::many(inputs))
        .statics(Statics::none())
        .sheets(Sheets::none())
        .formula_deps(FormulaCellDeps::none())
        .eval(move |ctx| {
            let mut sum = 0.0;
            for row in start_row..=end_row {
                let cell_addr = CellAddr::new(row, col);
                if let Some(n) = ctx.cell(cell_addr).as_number() {
                    // ROUND to 2 decimal places
                    sum += (n * 100.0).round() / 100.0;
                }
            }
            CellValue::Number(sum)
        })
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

/// Registriert eine Right Panel Calc Formel: =IF(M{row}="","",N{row}/M{row})
fn register_right_panel_calc(
    registry: &mut CellRegistry<Box<dyn Fn(&EvalContext) -> CellValue>>,
    addr: CellAddr,
    amount1_col: u16,
    amount2_col: u16,
) -> Result<(), RegistryError> {
    let row = addr.row;
    let excel_row = row + 1;
    let col1_letter = col_to_letter(amount1_col);
    let col2_letter = col_to_letter(amount2_col);

    let excel = Box::leak(
        format!(
            r#"=IF({}{}="","",{}{}/{}{})"#,
            col1_letter, excel_row, col2_letter, excel_row, col1_letter, excel_row
        )
        .into_boxed_str(),
    );

    let amount1_addr = CellAddr::new(row, amount1_col);
    let amount2_addr = CellAddr::new(row, amount2_col);

    let (addr, formula) = FormulaBuilder::new(addr, excel)
        .inputs(Inputs::many(vec![amount1_addr, amount2_addr]))
        .statics(Statics::none())
        .sheets(Sheets::none())
        .formula_deps(FormulaCellDeps::none())
        .eval(move |ctx| {
            let amount1 = ctx.cell(amount1_addr).as_number();
            let amount2 = ctx.cell(amount2_addr).as_number();

            match (amount1, amount2) {
                (Some(a1), Some(a2)) if a1 != 0.0 => CellValue::Number(a2 / a1),
                _ => CellValue::Empty,
            }
        })
        .build();

    registry.register_formula(addr, wrap_formula(formula))
}

// ============================================================================
// Evaluation Functions
// ============================================================================

/// Findet den Sprach-Index in TEXT_MATRIX
fn find_language_index(language: &str) -> Option<usize> {
    TEXT_MATRIX
        .iter()
        .position(|row| !row.is_empty() && row[0].eq_ignore_ascii_case(language))
}

/// Evaluiert: =IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,index,FALSE))
fn evaluate_text_lookup(ctx: &EvalContext, index: usize) -> CellValue {
    let language = match ctx.cell(addr::E2).as_text() {
        Some(lang) if !lang.is_empty() => lang,
        _ => return CellValue::Empty,
    };

    let lang_idx = match find_language_index(language) {
        Some(idx) => idx,
        None => return CellValue::Empty,
    };

    // Index ist 1-basiert in Excel, aber 0-basiert in TEXT_MATRIX
    let text_idx = index.saturating_sub(1);

    TEXT_MATRIX
        .get(lang_idx)
        .and_then(|row| row.get(text_idx))
        .map(|s| CellValue::Text(s.to_string()))
        .unwrap_or(CellValue::Empty)
}

/// Evaluiert mit Default-Wert
fn evaluate_text_lookup_default(ctx: &EvalContext, index: usize, default: &str) -> CellValue {
    let language = match ctx.cell(addr::E2).as_text() {
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

/// Evaluiert: =IF(E3="",VLOOKUP(...),E3)
fn evaluate_currency_or_lookup(ctx: &EvalContext, index: usize) -> CellValue {
    if let Some(currency) = ctx.cell(addr::E3).as_text() {
        if !currency.is_empty() {
            return CellValue::Text(currency.to_string());
        }
    }
    evaluate_text_lookup(ctx, index)
}

// ============================================================================
// Helpers
// ============================================================================

fn col_to_letter(col: u16) -> String {
    let mut result = String::new();
    let mut c = col as u32;
    loop {
        result.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    result
}

/// Wraps a typed FormulaCell into a boxed dynamic version
fn wrap_formula<E>(formula: FormulaCell<E>) -> FormulaCell<Box<dyn Fn(&EvalContext) -> CellValue>>
where
    E: Fn(&EvalContext) -> CellValue + 'static,
{
    FormulaCell {
        excel: formula.excel,
        deps: formula.deps,
        eval: Box::new(formula.eval),
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_build_registry() {
        let registry = build_registry().expect("Failed to build registry");

        // Prüfe dass API-Zellen registriert sind
        assert!(registry.is_api(addr::E2));
        assert!(registry.is_api(addr::E3));

        // Prüfe dass Formeln registriert sind
        assert!(registry.is_formula(CellAddr::new(0, 1))); // B1

        // Prüfe Anzahl
        assert!(registry.len() > 50);
    }

    #[test]
    fn test_no_overlaps() {
        let registry = build_registry().expect("Failed to build registry");

        // API und Formel dürfen sich nicht überschneiden
        for addr in registry.api_cells() {
            assert!(
                !registry.is_formula(*addr),
                "API cell {} is also a formula",
                addr
            );
        }
    }
}
