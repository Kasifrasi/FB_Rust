//! Dynamische Formel-Registrierung für Body- und Footer-Bereich
//!
//! Registriert alle dynamischen Formeln mit der einheitlichen CellRegistry.
//! Diese Formeln werden zur Runtime basierend auf dem BodyLayout/FooterLayout registriert.
//!
//! ## Body - Registrierte Formel-Typen
//!
//! ### VLOOKUP-Labels
//! - Kategorie-Labels (Header-Zeilen)
//! - Kategorie-Labels bei Header-Input (C-Spalte)
//! - Sum-Labels (Footer-Zeilen, B:C merged)
//! - Gesamt-Label (Total-Zeile, B:C merged)
//!
//! ### Ratio-Formeln (G-Spalte)
//! - IFERROR(F/D, 0) für alle Positions-Zeilen
//! - IFERROR(F/D, 0) für alle Header-Input-Zeilen
//! - IFERROR(F/D, 0) für alle Footer-Zeilen
//! - IFERROR(F/D, 0) für die Total-Zeile
//!
//! ### SUM-Formeln (Footer D, E, F)
//! - SUM(Dx:Dy) für jede Kategorie
//!
//! ### SUM-Formeln (Total D, E, F)
//! - SUM(D_footer1 + D_footer2 + ...) für jede Spalte
//!
//! ## Footer - Registrierte Formel-Typen
//!
//! ### VLOOKUP-Labels
//! - E(s):E(s+1): VLOOKUP(44) "Saldo für den Berichtszeitraum"
//! - B:D(s+1): VLOOKUP(43) "ABSCHLUSS"
//! - E(s+2): VLOOKUP(10) Währung
//! - B(s+4): VLOOKUP(45) "Saldo..."
//! - B(s+6): VLOOKUP(46) "Saldenabstimmung:"
//! - B(s+7-9): VLOOKUP(47,48,49) Bank/Kasse/Sonstiges
//! - B(s+13): VLOOKUP(50) Bestätigung 1
//! - B(s+14): VLOOKUP(54) Bestätigung 2
//! - B(s+19): VLOOKUP(51) "Ort, Datum..."
//! - D(s+19): VLOOKUP(52) "Unterschrift..."
//! - D(s+20): VLOOKUP(53) "Funktion..."
//!
//! ### Check-Formeln
//! - D(s+4): Check ✓ - IF(ROUND(E,2)=(ROUND(F-F,2)),"✓","")
//! - E(s+4): Differenz = E_income - E_total
//! - E(s+6): OK-Check - IF(E_saldo=SUM(inputs),"OK","")

use super::layout::FooterLayout;
use super::layout::{BodyLayout, CategoryMode, TOTAL_LABEL_INDEX};
use crate::report::api::{ApiKey, CellValue, FooterField, PositionField};
use crate::report::core::{
    CellAddr, DynRegistry, EvalContext, FormulaCell, FormulaCellDeps, FormulaDeps, Inputs,
    RegistryError, Sheets, Statics,
};

/// Registriert alle Body-Formeln basierend auf dem Layout
///
/// Diese Funktion sollte nach `build_registry()` und `BodyLayout::compute()`
/// aufgerufen werden, um alle dynamischen Formeln zu registrieren.
///
/// # Arguments
/// * `registry` - Die CellRegistry (muss bereits statische Zellen enthalten)
/// * `layout` - Das berechnete BodyLayout
///
/// # Registrierungsreihenfolge
/// 1. API-Zellen für alle Positions-Felder
/// 2. VLOOKUP-Formeln für Labels
/// 3. Ratio-Formeln (G-Spalte)
/// 4. SUM-Formeln (Footer)
/// 5. SUM-Formeln (Total)
pub fn register_body_formulas(
    registry: &mut DynRegistry,
    layout: &BodyLayout,
) -> Result<(), RegistryError> {
    // 1. API-Zellen für alle Positions-Felder
    register_body_api_cells(registry, layout)?;

    // 2. VLOOKUP-Formeln für Labels
    register_label_formulas(registry, layout)?;

    // 3. Ratio-Formeln (G-Spalte) - NACH den API-Zellen!
    register_ratio_formulas(registry, layout)?;

    // 4. SUM-Formeln (Footer) - NACH den API-Zellen!
    register_sum_formulas_footer(registry, layout)?;

    // 5. SUM-Formeln (Total) - NACH den Footer-SUM-Formeln!
    register_total_formulas(registry, layout)?;

    Ok(())
}

// ============================================================================
// API-Zellen Registrierung
// ============================================================================

/// Registriert alle Body-API-Zellen basierend auf dem Layout
fn register_body_api_cells(
    registry: &mut DynRegistry,
    layout: &BodyLayout,
) -> Result<(), RegistryError> {
    let mut batch: Vec<(ApiKey, CellAddr)> = Vec::new();

    for cat in &layout.categories {
        match &cat.mode {
            CategoryMode::HeaderInput { row } => {
                // Header-Input: D, E, F, H (nicht C - ist VLOOKUP Label)
                for field in PositionField::header_input_fields() {
                    let key = ApiKey::Position {
                        category: cat.meta.num,
                        position: 0,
                        field,
                    };
                    let addr = CellAddr::new(*row, field.col());
                    batch.push((key, addr));
                }
            }
            CategoryMode::WithPositions { positions, .. } => {
                // Positions-Modus: C, D, E, F, H für jede Position
                for i in 0..positions.count {
                    let row = positions.start_row + i as u32;
                    let pos = i + 1; // 1-basiert

                    for field in PositionField::all() {
                        let key = ApiKey::Position {
                            category: cat.meta.num,
                            position: pos,
                            field,
                        };
                        let addr = CellAddr::new(row, field.col());
                        batch.push((key, addr));
                    }
                }
            }
        }
    }

    registry.register_dynamic_api_batch(batch)
}

// ============================================================================
// VLOOKUP Label-Formeln
// ============================================================================

/// Registriert alle VLOOKUP-Formeln für Labels
fn register_label_formulas(
    registry: &mut DynRegistry,
    layout: &BodyLayout,
) -> Result<(), RegistryError> {
    for cat in &layout.categories {
        match &cat.mode {
            CategoryMode::HeaderInput { row } => {
                // C-Spalte: VLOOKUP für Kategorie-Label
                register_vlookup_formula(registry, CellAddr::new(*row, 2), cat.meta.label_index)?;
            }
            CategoryMode::WithPositions {
                header_row,
                footer_row,
                ..
            } => {
                // Header: C-Spalte mit VLOOKUP
                register_vlookup_formula(
                    registry,
                    CellAddr::new(*header_row, 2),
                    cat.meta.label_index,
                )?;

                // Footer: B-Spalte mit VLOOKUP (B:C wird gemerged beim Schreiben)
                register_vlookup_formula(
                    registry,
                    CellAddr::new(*footer_row, 1),
                    cat.meta.sum_label_index,
                )?;
            }
        }
    }

    // Total: B-Spalte mit VLOOKUP für "Gesamt"
    register_vlookup_formula(
        registry,
        CellAddr::new(layout.total_row, 1),
        TOTAL_LABEL_INDEX,
    )?;

    Ok(())
}

/// Registriert eine einzelne VLOOKUP-Formel
fn register_vlookup_formula(
    registry: &mut DynRegistry,
    addr: CellAddr,
    index: usize,
) -> Result<(), RegistryError> {
    let excel = Box::leak(
        format!(
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$CD,{},FALSE))"#,
            index
        )
        .into_boxed_str(),
    );

    let eval_fn: Box<dyn Fn(&EvalContext) -> CellValue> =
        Box::new(move |ctx: &EvalContext| ctx.vlookup_text(index));

    let formula = FormulaCell {
        excel,
        deps: FormulaDeps {
            inputs: Inputs::one(CellAddr::new(1, 4)), // E2
            statics: Statics::index(index),
            sheets: Sheets::lang_table(),
            formula_deps: FormulaCellDeps::none(),
        },
        eval: eval_fn,
    };

    registry.register_dynamic_formula(addr, formula)
}

// ============================================================================
// Ratio-Formeln (G-Spalte)
// ============================================================================

/// Registriert alle Ratio-Formeln (G-Spalte)
fn register_ratio_formulas(
    registry: &mut DynRegistry,
    layout: &BodyLayout,
) -> Result<(), RegistryError> {
    for row in layout.ratio_rows() {
        register_ratio_formula(registry, row)?;
    }
    Ok(())
}

/// Registriert eine einzelne Ratio-Formel
///
/// Formel: =IFERROR(F{row}/D{row},0)
fn register_ratio_formula(registry: &mut DynRegistry, row: u32) -> Result<(), RegistryError> {
    let addr = CellAddr::new(row, 6); // G-Spalte
    let d_addr = CellAddr::new(row, 3); // D-Spalte
    let f_addr = CellAddr::new(row, 5); // F-Spalte

    // Excel-Formel mit direkter Zellreferenz (1-basiert für Excel)
    let excel_row = row + 1;
    let excel: &'static str =
        Box::leak(format!("=IFERROR(F{}/D{},0)", excel_row, excel_row).into_boxed_str());

    let eval_fn: Box<dyn Fn(&EvalContext) -> CellValue> = Box::new(move |ctx: &EvalContext| {
        let d = ctx.cell(d_addr).as_number().unwrap_or(0.0);
        let f = ctx.cell(f_addr).as_number().unwrap_or(0.0);
        CellValue::Number(EvalContext::safe_divide(f, d))
    });

    let formula = FormulaCell {
        excel,
        deps: FormulaDeps {
            inputs: Inputs::none(),
            statics: Statics::none(),
            sheets: Sheets::none(),
            // D und F können API-Zellen (Positions-Zeilen) oder SUM-Formeln
            // (Footer/Total-Zeilen) sein. Topo-sort ignoriert nicht-Formel-Deps.
            formula_deps: FormulaCellDeps::many(vec![d_addr, f_addr]),
        },
        eval: eval_fn,
    };

    registry.register_dynamic_formula(addr, formula)
}

// ============================================================================
// SUM-Formeln (Footer D, E, F)
// ============================================================================

/// Registriert alle SUM-Formeln für Footer-Zeilen
fn register_sum_formulas_footer(
    registry: &mut DynRegistry,
    layout: &BodyLayout,
) -> Result<(), RegistryError> {
    for cat in &layout.categories {
        if let CategoryMode::WithPositions {
            positions,
            footer_row,
            ..
        } = &cat.mode
        {
            // D, E, F Spalten
            for col in [3u16, 4, 5] {
                register_sum_formula_range(
                    registry,
                    CellAddr::new(*footer_row, col),
                    positions.start_row,
                    positions.end_row,
                    col,
                )?;
            }
        }
    }
    Ok(())
}

/// Registriert eine einzelne SUM-Formel für einen Bereich
///
/// Formel: =SUM(Dx:Dy)
fn register_sum_formula_range(
    registry: &mut DynRegistry,
    addr: CellAddr,
    start_row: u32,
    end_row: u32,
    col: u16,
) -> Result<(), RegistryError> {
    let col_letter = CellAddr::col_to_letter(col);
    let excel = Box::leak(
        format!(
            "=SUM({}{}:{}{})",
            col_letter,
            start_row + 1, // Excel ist 1-basiert
            col_letter,
            end_row + 1
        )
        .into_boxed_str(),
    );

    // Sammle alle Input-Adressen
    let inputs: Vec<CellAddr> = (start_row..=end_row)
        .map(|r| CellAddr::new(r, col))
        .collect();

    let eval_fn: Box<dyn Fn(&EvalContext) -> CellValue> = Box::new(move |ctx: &EvalContext| {
        let mut sum = 0.0;
        for row in start_row..=end_row {
            let cell_addr = CellAddr::new(row, col);
            if let Some(n) = ctx.cell(cell_addr).as_number() {
                sum += n;
            }
        }
        CellValue::Number(sum)
    });

    let formula = FormulaCell {
        excel,
        deps: FormulaDeps {
            inputs: Inputs::many(inputs),
            statics: Statics::none(),
            sheets: Sheets::none(),
            formula_deps: FormulaCellDeps::none(),
        },
        eval: eval_fn,
    };

    registry.register_dynamic_formula(addr, formula)
}

// ============================================================================
// SUM-Formeln (Total D, E, F)
// ============================================================================

/// Registriert alle SUM-Formeln für die Total-Zeile
fn register_total_formulas(
    registry: &mut DynRegistry,
    layout: &BodyLayout,
) -> Result<(), RegistryError> {
    // Sammle alle "Summen-Zeilen" (Footer bei Positions-Modus, Header bei Header-Input)
    let sum_rows: Vec<u32> = layout.categories.iter().map(|c| c.sum_row()).collect();

    // D, E, F Spalten
    for col in [3u16, 4, 5] {
        register_sum_formula(
            registry,
            CellAddr::new(layout.total_row, col),
            &sum_rows,
            col,
        )?;
    }

    Ok(())
}

/// Registriert eine einzelne SUM-Formel
///
/// Formel: =SUM(D47+D69+...) - Summe aller Kategorie-Summen
fn register_sum_formula(
    registry: &mut DynRegistry,
    addr: CellAddr,
    sum_rows: &[u32],
    col: u16,
) -> Result<(), RegistryError> {
    let col_letter = CellAddr::col_to_letter(col);
    let refs: Vec<String> = sum_rows
        .iter()
        .map(|r| format!("{}{}", col_letter, r + 1))
        .collect();
    let excel = Box::leak(format!("=SUM({})", refs.join("+")).into_boxed_str());

    // Sammle Adressen (Footer-SUM-Formeln + Header-Input API-Zellen)
    let dep_addrs: Vec<CellAddr> = sum_rows.iter().map(|r| CellAddr::new(*r, col)).collect();
    let eval_addrs = dep_addrs.clone();

    let eval_fn: Box<dyn Fn(&EvalContext) -> CellValue> = Box::new(move |ctx: &EvalContext| {
        let sum: f64 = eval_addrs
            .iter()
            .filter_map(|a| ctx.cell(*a).as_number())
            .sum();
        CellValue::Number(sum)
    });

    let formula = FormulaCell {
        excel,
        deps: FormulaDeps {
            inputs: Inputs::none(),
            statics: Statics::none(),
            sheets: Sheets::none(),
            // Footer-SUM-Formeln werden als formula_deps erkannt,
            // Header-Input API-Zellen werden vom topo-sort ignoriert (korrekt).
            formula_deps: FormulaCellDeps::many(dep_addrs),
        },
        eval: eval_fn,
    };

    registry.register_dynamic_formula(addr, formula)
}

// ============================================================================
// Hilfsfunktionen
// ============================================================================

// ============================================================================
// Footer-Formeln Registrierung
// ============================================================================

/// Registriert alle Footer-Formeln basierend auf dem FooterLayout
///
/// Diese Funktion sollte nach `register_body_formulas()` aufgerufen werden.
///
/// # Arguments
/// * `registry` - Die CellRegistry
/// * `footer_layout` - Das berechnete FooterLayout
/// * `income_row` - Die Einnahmen-Zeile (0-indexiert, typischerweise 19 für Zeile 20)
///
/// # Registrierungsreihenfolge
/// 1. Footer API-Zellen (Bank, Kasse, Sonstiges)
/// 2. VLOOKUP-Formeln für Labels
/// 3. Check-Formeln (Differenz, OK)
pub fn register_footer_formulas(
    registry: &mut DynRegistry,
    footer_layout: &FooterLayout,
    income_row: u32,
) -> Result<(), RegistryError> {
    // 1. Footer API-Zellen registrieren
    register_footer_api_cells(registry, footer_layout)?;

    // 2. VLOOKUP-Formeln für alle Labels
    register_footer_label_formulas(registry, footer_layout)?;

    // 3. Check- und Differenz-Formeln
    register_footer_check_formulas(registry, footer_layout, income_row)?;

    Ok(())
}

/// Registriert Footer-API-Zellen (Bank, Kasse, Sonstiges)
fn register_footer_api_cells(
    registry: &mut DynRegistry,
    layout: &FooterLayout,
) -> Result<(), RegistryError> {
    let batch: Vec<(ApiKey, CellAddr)> = FooterField::all()
        .iter()
        .map(|field| {
            let addr = CellAddr::new(layout.input_rows[field.index()], field.col());
            (ApiKey::Footer(*field), addr)
        })
        .collect();

    registry.register_dynamic_api_batch(batch)
}

/// Registriert alle Footer VLOOKUP-Formeln
fn register_footer_label_formulas(
    registry: &mut DynRegistry,
    layout: &FooterLayout,
) -> Result<(), RegistryError> {
    let s = layout.start_row;

    // E(s):E(s+1) merged - VLOOKUP(44) "Saldo für den Berichtszeitraum"
    // Nur erste Zelle des Merge registrieren
    register_vlookup_formula(registry, CellAddr::new(s, 4), 44)?;

    // B:D(s+1) merged - VLOOKUP(43) "ABSCHLUSS"
    register_vlookup_formula(registry, CellAddr::new(s + 1, 1), 43)?;

    // E(s+2) - VLOOKUP(10) Währung
    register_vlookup_formula(registry, CellAddr::new(s + 2, 4), 10)?;

    // B(s+4) - VLOOKUP(45) "Saldo..."
    register_vlookup_formula(registry, CellAddr::new(s + 4, 1), 45)?;

    // B(s+6) - VLOOKUP(46) "Saldenabstimmung:"
    register_vlookup_formula(registry, CellAddr::new(s + 6, 1), 46)?;

    // B(s+7) - VLOOKUP(47) "Bank"
    register_vlookup_formula(registry, CellAddr::new(s + 7, 1), 47)?;

    // B(s+8) - VLOOKUP(48) "Kasse"
    register_vlookup_formula(registry, CellAddr::new(s + 8, 1), 48)?;

    // B(s+9) - VLOOKUP(49) "Sonstiges"
    register_vlookup_formula(registry, CellAddr::new(s + 9, 1), 49)?;

    // B(s+13) - VLOOKUP(50) Bestätigung 1
    register_vlookup_formula(registry, CellAddr::new(s + 13, 1), 50)?;

    // B(s+14) - VLOOKUP(54) Bestätigung 2
    register_vlookup_formula(registry, CellAddr::new(s + 14, 1), 54)?;

    // B(s+19) - VLOOKUP(51) "Ort, Datum..."
    register_vlookup_formula(registry, CellAddr::new(s + 19, 1), 51)?;

    // D(s+19) - VLOOKUP(52) "Unterschrift..."
    register_vlookup_formula(registry, CellAddr::new(s + 19, 3), 52)?;

    // D(s+20) - VLOOKUP(53) "Funktion..."
    register_vlookup_formula(registry, CellAddr::new(s + 20, 3), 53)?;

    Ok(())
}

/// Registriert Footer Check- und Differenz-Formeln
fn register_footer_check_formulas(
    registry: &mut DynRegistry,
    layout: &FooterLayout,
    income_row: u32,
) -> Result<(), RegistryError> {
    let s = layout.start_row;
    let total_row = layout.start_row - 3; // Footer startet 3 Zeilen nach Total

    // D(s+4): Check ✓ - IF(ROUND(E_saldo,2)=(ROUND(F_income-F_total,2)),"✓","")
    register_check_formula(registry, s, income_row, total_row)?;

    // E(s+4): Differenz = E_income - E_total
    register_diff_formula(registry, s, income_row, total_row)?;

    // E(s+6): OK-Check - IF(E_saldo=SUM(E_bank:E_sonstiges),"OK","")
    register_ok_formula(registry, layout)?;

    Ok(())
}

/// Registriert die Check-Formel (✓)
fn register_check_formula(
    registry: &mut DynRegistry,
    s: u32,
    income_row: u32,
    total_row: u32,
) -> Result<(), RegistryError> {
    let addr = CellAddr::new(s + 4, 3); // D(s+4)

    // Excel-Zeilen sind 1-basiert
    let saldo_row_excel = s + 4 + 1;
    let income_row_excel = income_row + 1;
    let total_row_excel = total_row + 1;

    let excel = Box::leak(
        format!(
            r#"=IF(ROUND(E{},2)=(ROUND(F{}-F{},2)),"✓","")"#,
            saldo_row_excel, income_row_excel, total_row_excel
        )
        .into_boxed_str(),
    );

    let f_income_addr = CellAddr::new(income_row, 5);
    let f_total_addr = CellAddr::new(total_row, 5);
    let e_saldo_addr = CellAddr::new(s + 4, 4); // E(s+4) - die Differenz-Formel

    let eval_fn: Box<dyn Fn(&EvalContext) -> CellValue> = Box::new(move |ctx: &EvalContext| {
        let e_saldo = ctx.cell(e_saldo_addr).as_number();
        let f_income = ctx.cell(f_income_addr).as_number();
        let f_total = ctx.cell(f_total_addr).as_number();

        match (e_saldo, f_income, f_total) {
            (Some(e_s), Some(f_inc), Some(f_tot)) => {
                let e_rounded = (e_s * 100.0).round();
                let f_diff_rounded = ((f_inc - f_tot) * 100.0).round();
                if e_rounded == f_diff_rounded {
                    CellValue::Text("✓".to_string())
                } else {
                    CellValue::Empty
                }
            }
            _ => CellValue::Empty,
        }
    });

    let formula = FormulaCell {
        excel,
        deps: FormulaDeps {
            inputs: Inputs::none(),
            statics: Statics::none(),
            sheets: Sheets::none(),
            // Abhängig von E_saldo (Diff-Formel), F_income und F_total (SUM-Formeln)
            formula_deps: FormulaCellDeps::many(vec![e_saldo_addr, f_income_addr, f_total_addr]),
        },
        eval: eval_fn,
    };

    registry.register_dynamic_formula(addr, formula)
}

/// Registriert die Differenz-Formel (E_income - E_total)
fn register_diff_formula(
    registry: &mut DynRegistry,
    s: u32,
    income_row: u32,
    total_row: u32,
) -> Result<(), RegistryError> {
    let addr = CellAddr::new(s + 4, 4); // E(s+4)

    let income_row_excel = income_row + 1;
    let total_row_excel = total_row + 1;

    let excel = Box::leak(format!("=E{}-E{}", income_row_excel, total_row_excel).into_boxed_str());

    let e_income_addr = CellAddr::new(income_row, 4);
    let e_total_addr = CellAddr::new(total_row, 4);

    let eval_fn: Box<dyn Fn(&EvalContext) -> CellValue> = Box::new(move |ctx: &EvalContext| {
        let e_income = ctx.cell(e_income_addr).as_number().unwrap_or(0.0);
        let e_total = ctx.cell(e_total_addr).as_number().unwrap_or(0.0);
        CellValue::Number(e_income - e_total)
    });

    let formula = FormulaCell {
        excel,
        deps: FormulaDeps {
            inputs: Inputs::none(), // E_income und E_total sind Formeln
            statics: Statics::none(),
            sheets: Sheets::none(),
            formula_deps: FormulaCellDeps::many(vec![e_income_addr, e_total_addr]),
        },
        eval: eval_fn,
    };

    registry.register_dynamic_formula(addr, formula)
}

/// Registriert die OK-Check Formel
fn register_ok_formula(
    registry: &mut DynRegistry,
    layout: &FooterLayout,
) -> Result<(), RegistryError> {
    let s = layout.start_row;
    let addr = CellAddr::new(s + 6, 4); // E(s+6)

    // Excel-Zeilen sind 1-basiert
    let saldo_row_excel = s + 4 + 1;
    let bank_row_excel = s + 7 + 1;
    let sonstiges_row_excel = s + 9 + 1;

    let excel = Box::leak(
        format!(
            r#"=IF(E{}=SUM(E{}:E{}),"OK","")"#,
            saldo_row_excel, bank_row_excel, sonstiges_row_excel
        )
        .into_boxed_str(),
    );

    let e_saldo_addr = CellAddr::new(s + 4, 4);
    let e_bank_addr = CellAddr::new(layout.input_rows[0], 4);
    let e_kasse_addr = CellAddr::new(layout.input_rows[1], 4);
    let e_sonstiges_addr = CellAddr::new(layout.input_rows[2], 4);

    let eval_fn: Box<dyn Fn(&EvalContext) -> CellValue> = Box::new(move |ctx: &EvalContext| {
        let e_saldo = ctx.cell(e_saldo_addr).as_number();
        let bank = ctx.cell(e_bank_addr).as_number().unwrap_or(0.0);
        let kasse = ctx.cell(e_kasse_addr).as_number().unwrap_or(0.0);
        let sonstiges = ctx.cell(e_sonstiges_addr).as_number().unwrap_or(0.0);

        match e_saldo {
            Some(saldo) => {
                let sum_inputs = bank + kasse + sonstiges;
                if (saldo * 100.0).round() == (sum_inputs * 100.0).round() {
                    CellValue::Text("OK".to_string())
                } else {
                    CellValue::Empty
                }
            }
            None => CellValue::Empty,
        }
    });

    let formula = FormulaCell {
        excel,
        deps: FormulaDeps {
            inputs: Inputs::many(vec![e_bank_addr, e_kasse_addr, e_sonstiges_addr]),
            statics: Statics::none(),
            sheets: Sheets::none(),
            formula_deps: FormulaCellDeps::many(vec![e_saldo_addr]),
        },
        eval: eval_fn,
    };

    registry.register_dynamic_formula(addr, formula)
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::report::body::config::BodyConfig;
    use crate::report::core::build_registry;

    #[test]
    fn test_register_body_formulas() {
        let mut registry = build_registry().expect("Failed to build registry");
        let config = BodyConfig::new()
            .with_positions(1, 2)
            .with_positions(2, 0) // Header-Input
            .with_positions(3, 3);
        let layout = BodyLayout::compute(&config);

        // Registriere Body-Formeln
        register_body_formulas(&mut registry, &layout).expect("Failed to register body formulas");

        // Prüfe dass dynamische Zellen registriert wurden
        assert!(registry.has_dynamic_cells());
        assert!(registry.dynamic_cell_count() > 0);

        // Prüfe API-Zellen für Positionen
        // Kategorie 1, Position 1, Approved (D-Spalte)
        let cat1_pos1_d = layout.position_addr(1, 1, PositionField::Approved).unwrap();
        assert!(registry.is_api(cat1_pos1_d));

        // Prüfe VLOOKUP-Formel für Kategorie-Label
        let cat1_header = CellAddr::new(26, 2); // C27
        assert!(registry.is_formula(cat1_header));

        // Prüfe Ratio-Formel
        let ratio_row = 27; // Erste Position von Kategorie 1
        let ratio_addr = CellAddr::new(ratio_row, 6); // G28
        assert!(registry.is_formula(ratio_addr));
    }

    #[test]
    fn test_header_input_mode() {
        let mut registry = build_registry().expect("Failed to build registry");
        let config = BodyConfig::new()
            .with_positions(1, 0) // Header-Input
            .with_positions(2, 0);
        let layout = BodyLayout::compute(&config);

        register_body_formulas(&mut registry, &layout).expect("Failed to register body formulas");

        // Kategorie 1 ist Header-Input
        // API-Zellen: D, E, F, H (nicht C!)
        let row = 26;
        assert!(registry.is_api(CellAddr::new(row, 3))); // D
        assert!(registry.is_api(CellAddr::new(row, 4))); // E
        assert!(registry.is_api(CellAddr::new(row, 5))); // F
        assert!(registry.is_api(CellAddr::new(row, 7))); // H

        // C ist VLOOKUP-Formel, keine API-Zelle
        assert!(registry.is_formula(CellAddr::new(row, 2))); // C
        assert!(!registry.is_api(CellAddr::new(row, 2)));
    }
}
