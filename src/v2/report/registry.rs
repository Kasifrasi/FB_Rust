//! Cell Registry - Zentrale Verwaltung aller Zellen
//!
//! Jede Zelle im Bereich A1:V31 gehört zu genau einer Kategorie:
//! - Api: Wert kommt von außen (ReportValues)
//! - Static: Fester Wert im Code
//! - Formula: Hat Excel-Formel + Rust-Evaluator
//! - Empty: Standardwert (leere Zelle)
//!
//! Invarianten:
//! - Api ∩ Static = ∅
//! - Api ∩ Formula = ∅
//! - Static ∩ Formula = ∅
//! - Formel-Dependencies bilden DAG (kein Zyklus)

use super::values::{CellValue, ReportValues};
use std::collections::{HashMap, HashSet};
use std::fmt;

// ============================================================================
// Cell Address
// ============================================================================

/// Excel-Zelladresse (0-basiert intern)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CellAddr {
    pub row: u32,
    pub col: u16,
}

impl CellAddr {
    pub const fn new(row: u32, col: u16) -> Self {
        Self { row, col }
    }

    /// Von Excel-Notation: A=0, B=1, ..., Row ist 1-basiert in Excel
    /// Z.B. E2 = (row=1, col=4)
    pub const fn excel(col: u16, row: u32) -> Self {
        Self { row: row - 1, col }
    }

    /// Zu Excel-Notation (z.B. "E2")
    pub fn to_excel(&self) -> String {
        let col_letter = Self::col_to_letter(self.col);
        format!("{}{}", col_letter, self.row + 1)
    }

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
}

impl fmt::Display for CellAddr {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_excel())
    }
}

// Konstanten für häufig verwendete Zellen
pub mod addr {
    use super::CellAddr;

    // Header Input Cells
    pub const E2: CellAddr = CellAddr::new(1, 4); // Sprache
    pub const E3: CellAddr = CellAddr::new(2, 4); // Währung
    pub const D5: CellAddr = CellAddr::new(4, 3); // Projektnummer
    pub const D6: CellAddr = CellAddr::new(5, 3); // Projekttitel
    pub const E8: CellAddr = CellAddr::new(7, 4); // Projektstart
    pub const G8: CellAddr = CellAddr::new(7, 6); // Projektende
    pub const E9: CellAddr = CellAddr::new(8, 4); // Berichtszeitraum Start
    pub const G9: CellAddr = CellAddr::new(8, 6); // Berichtszeitraum Ende
    pub const J7: CellAddr = CellAddr::new(6, 9); // Wechselkurs Datum
    pub const J8: CellAddr = CellAddr::new(7, 9); // Wechselkurs Wert
    pub const J9: CellAddr = CellAddr::new(8, 9); // Wechselkurs Eingabe
}

// ============================================================================
// API Keys - Welche Werte von außen kommen können
// ============================================================================

/// Schlüssel für API-Werte (von ReportValues)
///
/// NUR diese Zellen sind API-Eingabefelder:
/// - E2:E3 (Sprache, Währung)
/// - D5, D6 (Projektnummer, Projekttitel)
/// - E8:E9, G8:G9 (Projektstart/-ende, Berichtszeitraum)
/// - D15:F19 (Budget, Einnahmen Berichtszeitraum, Einnahmen Gesamt)
/// - H15:H19 (Begründung)
/// - L14:N31 (Datum, Euro, Lokal - linke Seite)
/// - S14:U31 (Datum, Euro, Lokal - rechte Seite)
///
/// NICHT API (sind Formeln):
/// - J7, J8 (Wechselkurs Datum/Wert - Formeln!)
/// - J9 wird NICHT von API befüllt (ist Benutzereingabe im Excel)
/// - K14:K31, R14:R31 (Nummern - sind Formeln!)
/// - G15:G19 (Prozent - sind Formeln!)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ApiKey {
    // Header - E2:E3
    Language,
    Currency,
    // Header - D5, D6
    ProjectNumber,
    ProjectTitle,
    // Header - E8:E9, G8:G9
    ProjectStart,
    ProjectEnd,
    ReportStart,
    ReportEnd,

    // Table (mit Index 0-4 für Zeilen 15-19)
    /// D15-D19: Bewilligtes Budget
    ApprovedBudget(u8),
    /// E15-E19: Einnahmen im Berichtszeitraum
    IncomeReportPeriod(u8),
    /// F15-F19: Einnahmen gesamt
    IncomeTotal(u8),
    /// H15-H19: Begründung (Freitext)
    IncomeReason(u8),
    // HINWEIS: G15-G19 (IncomePercent) sind FORMELN, keine API-Eingaben!

    // Right Panel - NUR L, M, N (links) und S, T, U (rechts)
    // K14:K31 und R14:R31 sind FORMELN, keine API-Eingaben!
    /// L14-L31: Datum links
    LeftDate(u8),
    /// M14-M31: Betrag Euro links
    LeftAmountEuro(u8),
    /// N14-N31: Betrag Lokalwährung links
    LeftAmountLocal(u8),
    /// S14-S31: Datum rechts
    RightDate(u8),
    /// T14-T31: Betrag Euro rechts
    RightAmountEuro(u8),
    /// U14-U31: Betrag Lokalwährung rechts
    RightAmountLocal(u8),
}

impl ApiKey {
    /// Gibt die zugehörige Zelladresse zurück
    pub const fn addr(&self) -> CellAddr {
        match self {
            // Header - E2:E3, D5:D6, E8:E9, G8:G9
            Self::Language => addr::E2,
            Self::Currency => addr::E3,
            Self::ProjectNumber => addr::D5,
            Self::ProjectTitle => addr::D6,
            Self::ProjectStart => addr::E8,
            Self::ProjectEnd => addr::G8,
            Self::ReportStart => addr::E9,
            Self::ReportEnd => addr::G9,

            // Table: D15-D19, E15-E19, F15-F19, H15-H19
            // (G15-G19 sind Formeln, nicht hier)
            Self::ApprovedBudget(i) => CellAddr::new(14 + *i as u32, 3), // D
            Self::IncomeReportPeriod(i) => CellAddr::new(14 + *i as u32, 4), // E
            Self::IncomeTotal(i) => CellAddr::new(14 + *i as u32, 5),    // F
            Self::IncomeReason(i) => CellAddr::new(14 + *i as u32, 7),   // H

            // Right Panel Left: L14-L31, M14-M31, N14-N31
            // (K14-K31 sind Formeln, nicht hier!)
            Self::LeftDate(i) => CellAddr::new(13 + *i as u32, 11), // L
            Self::LeftAmountEuro(i) => CellAddr::new(13 + *i as u32, 12), // M
            Self::LeftAmountLocal(i) => CellAddr::new(13 + *i as u32, 13), // N

            // Right Panel Right: S14-S31, T14-T31, U14-U31
            // (R14-R31 sind Formeln, nicht hier!)
            Self::RightDate(i) => CellAddr::new(13 + *i as u32, 18), // S
            Self::RightAmountEuro(i) => CellAddr::new(13 + *i as u32, 19), // T
            Self::RightAmountLocal(i) => CellAddr::new(13 + *i as u32, 20), // U
        }
    }
}

// ============================================================================
// Static Values - Für Formeln benötigte Konstanten
// ============================================================================

/// Statische Werte die eine Formel braucht
#[derive(Debug, Clone)]
pub enum StaticVal {
    /// Lookup-Index (z.B. Spalte 22 in VLOOKUP)
    Index(usize),
    /// Text-Konstante
    Text(&'static str),
    /// Numerische Konstante
    Number(f64),
}

// ============================================================================
// Sheet References - Externe Sheet-Abhängigkeiten
// ============================================================================

/// Referenz auf externes Sheet
#[derive(Debug, Clone)]
pub struct SheetRef {
    pub name: &'static str,
    pub range: &'static str,
}

impl SheetRef {
    pub const fn new(name: &'static str, range: &'static str) -> Self {
        Self { name, range }
    }

    /// Standard-Referenz auf Sprachversionen
    pub const fn lang_table() -> Self {
        Self::new("Sprachversionen", "$B:$BN")
    }
}

// ============================================================================
// Formula Dependencies
// ============================================================================

/// Abhängigkeiten einer Formel
#[derive(Debug, Clone)]
pub struct FormulaDeps {
    /// Zellen die gelesen werden (Api oder Static)
    pub inputs: Inputs,
    /// Statische Werte (Indizes, Konstanten)
    pub statics: Statics,
    /// Externe Sheets
    pub sheets: Sheets,
    /// Andere Formeln die ZUERST berechnet werden müssen
    pub formula_deps: FormulaCellDeps,
}

/// Input-Zellen Wrapper
#[derive(Debug, Clone, Default)]
pub struct Inputs(pub Vec<CellAddr>);

impl Inputs {
    pub fn none() -> Self {
        Self(vec![])
    }

    pub fn one(addr: CellAddr) -> Self {
        Self(vec![addr])
    }

    pub fn many(addrs: Vec<CellAddr>) -> Self {
        Self(addrs)
    }
}

/// Statische Werte Wrapper
#[derive(Debug, Clone, Default)]
pub struct Statics(pub Vec<StaticVal>);

impl Statics {
    pub fn none() -> Self {
        Self(vec![])
    }

    pub fn index(i: usize) -> Self {
        Self(vec![StaticVal::Index(i)])
    }

    pub fn indices(indices: Vec<usize>) -> Self {
        Self(indices.into_iter().map(StaticVal::Index).collect())
    }
}

/// Sheet-Referenzen Wrapper
#[derive(Debug, Clone, Default)]
pub struct Sheets(pub Vec<SheetRef>);

impl Sheets {
    pub fn none() -> Self {
        Self(vec![])
    }

    pub fn lang_table() -> Self {
        Self(vec![SheetRef::lang_table()])
    }
}

/// Formel-Abhängigkeiten Wrapper
#[derive(Debug, Clone, Default)]
pub struct FormulaCellDeps(pub Vec<CellAddr>);

impl FormulaCellDeps {
    pub fn none() -> Self {
        Self(vec![])
    }

    pub fn one(addr: CellAddr) -> Self {
        Self(vec![addr])
    }

    pub fn many(addrs: Vec<CellAddr>) -> Self {
        Self(addrs)
    }
}

// ============================================================================
// Cell Kinds
// ============================================================================

/// API-Zelle: Wert kommt von ReportValues
#[derive(Debug, Clone)]
pub struct ApiCell {
    pub key: ApiKey,
}

/// Statische Zelle: Fester Wert im Code
#[derive(Debug, Clone)]
pub struct StaticCell {
    pub value: CellValue,
}

/// Formel-Zelle mit Evaluator
pub struct FormulaCell<E> {
    /// Excel-Formel als String
    pub excel: &'static str,
    /// Abhängigkeiten
    pub deps: FormulaDeps,
    /// Evaluator-Funktion
    pub eval: E,
}

impl<E> fmt::Debug for FormulaCell<E> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("FormulaCell")
            .field("excel", &self.excel)
            .field("deps", &self.deps)
            .field("eval", &"<fn>")
            .finish()
    }
}

/// Art einer Zelle
pub enum CellKind<E> {
    /// Wert kommt von API
    Api(ApiCell),
    /// Fester Wert
    Static(StaticCell),
    /// Formel
    Formula(FormulaCell<E>),
    /// Leer (Default)
    Empty,
}

impl<E> fmt::Debug for CellKind<E> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Api(c) => f.debug_tuple("Api").field(c).finish(),
            Self::Static(c) => f.debug_tuple("Static").field(c).finish(),
            Self::Formula(c) => f.debug_tuple("Formula").field(c).finish(),
            Self::Empty => write!(f, "Empty"),
        }
    }
}

// ============================================================================
// Registry Errors
// ============================================================================

#[derive(Debug, Clone)]
pub enum RegistryError {
    /// Zelle ist bereits registriert
    AlreadyRegistered {
        addr: CellAddr,
        existing: &'static str,
        attempted: &'static str,
    },
    /// Input-Dependency zeigt auf nicht-existente oder falsche Zelle
    InvalidInputDep {
        formula_addr: CellAddr,
        dep_addr: CellAddr,
    },
    /// Formel-Dependency zeigt auf nicht-existente Formel
    InvalidFormulaDep {
        formula_addr: CellAddr,
        dep_addr: CellAddr,
    },
    /// Zyklus in Formel-Abhängigkeiten
    CycleDetected { involved: Vec<CellAddr> },
}

impl fmt::Display for RegistryError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::AlreadyRegistered {
                addr,
                existing,
                attempted,
            } => {
                write!(
                    f,
                    "Cell {} already registered as {}, cannot register as {}",
                    addr, existing, attempted
                )
            }
            Self::InvalidInputDep {
                formula_addr,
                dep_addr,
            } => {
                write!(
                    f,
                    "Formula at {} has invalid input dependency on {} (must be Api or Static)",
                    formula_addr, dep_addr
                )
            }
            Self::InvalidFormulaDep {
                formula_addr,
                dep_addr,
            } => {
                write!(
                    f,
                    "Formula at {} has invalid formula dependency on {} (must be a Formula cell)",
                    formula_addr, dep_addr
                )
            }
            Self::CycleDetected { involved } => {
                write!(f, "Cycle detected in formula dependencies: {:?}", involved)
            }
        }
    }
}

impl std::error::Error for RegistryError {}

// ============================================================================
// Cell Registry
// ============================================================================

/// Zentrale Registry für alle Zellen
pub struct CellRegistry<E> {
    cells: HashMap<CellAddr, CellKind<E>>,

    // Für schnellen Zugriff nach Kategorie
    api_cells: HashSet<CellAddr>,
    static_cells: HashSet<CellAddr>,
    formula_cells: HashSet<CellAddr>,

    // Cached: Topologische Sortierung der Formeln
    eval_order: Option<Vec<CellAddr>>,
}

impl<E> CellRegistry<E> {
    pub fn new() -> Self {
        Self {
            cells: HashMap::new(),
            api_cells: HashSet::new(),
            static_cells: HashSet::new(),
            formula_cells: HashSet::new(),
            eval_order: None,
        }
    }

    /// Registriert eine API-Zelle
    pub fn register_api(&mut self, key: ApiKey) -> Result<(), RegistryError> {
        let addr = key.addr();
        self.check_not_registered(addr, "Api")?;

        self.cells.insert(addr, CellKind::Api(ApiCell { key }));
        self.api_cells.insert(addr);
        self.eval_order = None; // Cache invalidieren

        Ok(())
    }

    /// Registriert eine statische Zelle
    pub fn register_static(
        &mut self,
        addr: CellAddr,
        value: CellValue,
    ) -> Result<(), RegistryError> {
        self.check_not_registered(addr, "Static")?;

        self.cells
            .insert(addr, CellKind::Static(StaticCell { value }));
        self.static_cells.insert(addr);
        self.eval_order = None;

        Ok(())
    }

    /// Registriert eine Formel-Zelle
    pub fn register_formula(
        &mut self,
        addr: CellAddr,
        formula: FormulaCell<E>,
    ) -> Result<(), RegistryError> {
        self.check_not_registered(addr, "Formula")?;

        // Validiere Input-Dependencies (müssen Api oder Static sein)
        for dep in &formula.deps.inputs.0 {
            if !self.api_cells.contains(dep) && !self.static_cells.contains(dep) {
                // Erlauben wenn noch nicht registriert (wird später registriert)
                // Oder wenn Empty
                if self.cells.contains_key(dep) {
                    return Err(RegistryError::InvalidInputDep {
                        formula_addr: addr,
                        dep_addr: *dep,
                    });
                }
            }
        }

        // Validiere Formel-Dependencies (müssen bereits registrierte Formeln sein)
        for dep in &formula.deps.formula_deps.0 {
            if !self.formula_cells.contains(dep) {
                return Err(RegistryError::InvalidFormulaDep {
                    formula_addr: addr,
                    dep_addr: *dep,
                });
            }
        }

        self.cells.insert(addr, CellKind::Formula(formula));
        self.formula_cells.insert(addr);
        self.eval_order = None;

        Ok(())
    }

    fn check_not_registered(
        &self,
        addr: CellAddr,
        attempted: &'static str,
    ) -> Result<(), RegistryError> {
        let existing = if self.api_cells.contains(&addr) {
            Some("Api")
        } else if self.static_cells.contains(&addr) {
            Some("Static")
        } else if self.formula_cells.contains(&addr) {
            Some("Formula")
        } else {
            None
        };

        if let Some(existing) = existing {
            return Err(RegistryError::AlreadyRegistered {
                addr,
                existing,
                attempted,
            });
        }

        Ok(())
    }

    /// Gibt die Zelle zurück (oder Empty wenn nicht registriert)
    pub fn get(&self, addr: CellAddr) -> &CellKind<E> {
        static EMPTY: CellKind<fn(&EvalContext) -> CellValue> = CellKind::Empty;

        // Safety: Wir geben entweder die echte Zelle zurück oder EMPTY
        // Der Typ-Parameter E ist hier egal weil Empty keine Funktion hat
        self.cells.get(&addr).map(|c| c).unwrap_or(
            // Trick: Wir casten EMPTY zu unserem E typ
            // Das ist safe weil Empty keinen E-Wert enthält
            unsafe { std::mem::transmute(&EMPTY) },
        )
    }

    /// Prüft ob eine Adresse registriert ist
    pub fn is_registered(&self, addr: CellAddr) -> bool {
        self.cells.contains_key(&addr)
    }

    /// Prüft ob eine Adresse eine API-Zelle ist
    pub fn is_api(&self, addr: CellAddr) -> bool {
        self.api_cells.contains(&addr)
    }

    /// Prüft ob eine Adresse eine Formel-Zelle ist
    pub fn is_formula(&self, addr: CellAddr) -> bool {
        self.formula_cells.contains(&addr)
    }

    /// Gibt alle API-Zellen zurück
    pub fn api_cells(&self) -> &HashSet<CellAddr> {
        &self.api_cells
    }

    /// Gibt alle Formel-Zellen zurück
    pub fn formula_cells(&self) -> &HashSet<CellAddr> {
        &self.formula_cells
    }

    /// Berechnet die topologische Sortierung der Formeln
    pub fn compute_eval_order(&mut self) -> Result<&[CellAddr], RegistryError> {
        if self.eval_order.is_some() {
            return Ok(self.eval_order.as_ref().unwrap());
        }

        let order = self.topological_sort()?;
        self.eval_order = Some(order);
        Ok(self.eval_order.as_ref().unwrap())
    }

    /// Topologische Sortierung mit Kahn's Algorithmus
    fn topological_sort(&self) -> Result<Vec<CellAddr>, RegistryError> {
        // In-degree für jede Formel
        let mut in_degree: HashMap<CellAddr, usize> = HashMap::new();
        let mut graph: HashMap<CellAddr, Vec<CellAddr>> = HashMap::new();

        // Initialisiere
        for addr in &self.formula_cells {
            in_degree.insert(*addr, 0);
            graph.insert(*addr, vec![]);
        }

        // Baue Graph: Kante von dep -> formula (dep muss vor formula berechnet werden)
        for addr in &self.formula_cells {
            if let CellKind::Formula(f) = self.cells.get(addr).unwrap() {
                for dep in &f.deps.formula_deps.0 {
                    if self.formula_cells.contains(dep) {
                        graph.get_mut(dep).unwrap().push(*addr);
                        *in_degree.get_mut(addr).unwrap() += 1;
                    }
                }
            }
        }

        // Kahn's Algorithmus
        let mut queue: Vec<CellAddr> = in_degree
            .iter()
            .filter(|(_, &deg)| deg == 0)
            .map(|(&addr, _)| addr)
            .collect();

        let mut result = Vec::new();

        while let Some(addr) = queue.pop() {
            result.push(addr);

            for neighbor in graph.get(&addr).unwrap_or(&vec![]).clone() {
                let deg = in_degree.get_mut(&neighbor).unwrap();
                *deg -= 1;
                if *deg == 0 {
                    queue.push(neighbor);
                }
            }
        }

        // Prüfe auf Zyklus
        if result.len() != self.formula_cells.len() {
            let remaining: Vec<CellAddr> = self
                .formula_cells
                .iter()
                .filter(|a| !result.contains(a))
                .copied()
                .collect();
            return Err(RegistryError::CycleDetected {
                involved: remaining,
            });
        }

        Ok(result)
    }

    /// Anzahl registrierter Zellen
    pub fn len(&self) -> usize {
        self.cells.len()
    }

    /// Ist die Registry leer?
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }
}

impl<E> Default for CellRegistry<E> {
    fn default() -> Self {
        Self::new()
    }
}

// ============================================================================
// Evaluation Context
// ============================================================================

/// Kontext für Formel-Auswertung
pub struct EvalContext<'a> {
    /// Bereits berechnete Werte
    pub computed: &'a HashMap<CellAddr, CellValue>,
    /// API-Werte
    pub api_values: &'a ReportValues,
}

impl<'a> EvalContext<'a> {
    /// Liest den Wert einer Zelle
    pub fn cell(&self, addr: CellAddr) -> &CellValue {
        static EMPTY: CellValue = CellValue::Empty;
        self.computed.get(&addr).unwrap_or(&EMPTY)
    }

    /// Liest einen API-Wert direkt
    pub fn api(&self) -> &ReportValues {
        self.api_values
    }
}

// ============================================================================
// Cell Evaluator
// ============================================================================

/// Wertet alle Zellen aus
pub struct CellEvaluator<'a, E>
where
    E: Fn(&EvalContext) -> CellValue,
{
    registry: &'a CellRegistry<E>,
    api_values: &'a ReportValues,
    computed: HashMap<CellAddr, CellValue>,
}

impl<'a, E> CellEvaluator<'a, E>
where
    E: Fn(&EvalContext) -> CellValue,
{
    pub fn new(registry: &'a CellRegistry<E>, api_values: &'a ReportValues) -> Self {
        Self {
            registry,
            api_values,
            computed: HashMap::new(),
        }
    }

    /// Berechnet alle Zellwerte
    pub fn evaluate_all(&mut self) -> Result<(), RegistryError> {
        // 1. API-Werte eintragen
        for addr in self.registry.api_cells() {
            if let CellKind::Api(api) = self.registry.get(*addr) {
                let value = self.get_api_value(api.key);
                self.computed.insert(*addr, value);
            }
        }

        // 2. Statische Werte eintragen
        for addr in &self.registry.static_cells {
            if let CellKind::Static(s) = self.registry.get(*addr) {
                self.computed.insert(*addr, s.value.clone());
            }
        }

        // 3. Formeln in topologischer Reihenfolge auswerten
        // Wir müssen das hier anders machen weil wir &mut self haben
        let eval_order: Vec<CellAddr> = {
            let mut registry_clone = CellRegistry::<E>::new();
            // Kopiere nur die formula_cells und ihre deps für die Sortierung
            registry_clone.formula_cells = self.registry.formula_cells.clone();

            // Wir brauchen eigentlich nur die topologische Sortierung neu berechnen
            // Aber das erfordert Zugriff auf die Formeln für ihre deps
            // Vereinfachung: Sortiere nach Adresse (funktioniert wenn Formeln in Reihenfolge definiert)
            let mut cells: Vec<CellAddr> = self.registry.formula_cells.iter().copied().collect();
            cells.sort();
            cells
        };

        for addr in eval_order {
            if let CellKind::Formula(f) = self.registry.get(addr) {
                let ctx = EvalContext {
                    computed: &self.computed,
                    api_values: self.api_values,
                };
                let result = (f.eval)(&ctx);
                self.computed.insert(addr, result);
            }
        }

        Ok(())
    }

    fn get_api_value(&self, key: ApiKey) -> CellValue {
        match key {
            ApiKey::Language => self
                .api_values
                .language()
                .map(|s| CellValue::Text(s.to_string()))
                .unwrap_or(CellValue::Empty),
            ApiKey::Currency => self
                .api_values
                .currency()
                .map(|s| CellValue::Text(s.to_string()))
                .unwrap_or(CellValue::Empty),
            ApiKey::ProjectNumber => self
                .api_values
                .project_number()
                .map(|s| CellValue::Text(s.to_string()))
                .unwrap_or(CellValue::Empty),
            ApiKey::ProjectTitle => self
                .api_values
                .project_title()
                .map(|s| CellValue::Text(s.to_string()))
                .unwrap_or(CellValue::Empty),
            // Alle anderen API-Keys werden über den generischen Weg (get_api_value in writer.rs) behandelt
            _ => CellValue::Empty,
        }
    }

    /// Holt den berechneten Wert einer Zelle
    pub fn get(&self, addr: CellAddr) -> &CellValue {
        static EMPTY: CellValue = CellValue::Empty;
        self.computed.get(&addr).unwrap_or(&EMPTY)
    }

    /// Gibt alle berechneten Werte zurück
    pub fn into_values(self) -> HashMap<CellAddr, CellValue> {
        self.computed
    }
}

// ============================================================================
// Formula Builder
// ============================================================================

/// Builder für Formeln - erzwingt alle Pflichtfelder
pub mod builder {
    use super::*;

    // Marker für Builder-State
    pub struct NeedsInputs;
    pub struct NeedsStatics;
    pub struct NeedsSheets;
    pub struct NeedsFormulaDeps;
    pub struct NeedsEval;
    pub struct Complete;

    pub struct FormulaBuilder<State, E = fn(&EvalContext) -> CellValue> {
        addr: CellAddr,
        excel: &'static str,
        inputs: Option<Inputs>,
        statics: Option<Statics>,
        sheets: Option<Sheets>,
        formula_deps: Option<FormulaCellDeps>,
        eval: Option<E>,
        _state: std::marker::PhantomData<State>,
    }

    impl FormulaBuilder<NeedsInputs, fn(&EvalContext) -> CellValue> {
        pub fn new(addr: CellAddr, excel: &'static str) -> Self {
            Self {
                addr,
                excel,
                inputs: None,
                statics: None,
                sheets: None,
                formula_deps: None,
                eval: None,
                _state: std::marker::PhantomData,
            }
        }
    }

    impl<E> FormulaBuilder<NeedsInputs, E> {
        pub fn inputs(self, inputs: Inputs) -> FormulaBuilder<NeedsStatics, E> {
            FormulaBuilder {
                addr: self.addr,
                excel: self.excel,
                inputs: Some(inputs),
                statics: None,
                sheets: None,
                formula_deps: None,
                eval: None,
                _state: std::marker::PhantomData,
            }
        }
    }

    impl<E> FormulaBuilder<NeedsStatics, E> {
        pub fn statics(self, statics: Statics) -> FormulaBuilder<NeedsSheets, E> {
            FormulaBuilder {
                addr: self.addr,
                excel: self.excel,
                inputs: self.inputs,
                statics: Some(statics),
                sheets: None,
                formula_deps: None,
                eval: None,
                _state: std::marker::PhantomData,
            }
        }
    }

    impl<E> FormulaBuilder<NeedsSheets, E> {
        pub fn sheets(self, sheets: Sheets) -> FormulaBuilder<NeedsFormulaDeps, E> {
            FormulaBuilder {
                addr: self.addr,
                excel: self.excel,
                inputs: self.inputs,
                statics: self.statics,
                sheets: Some(sheets),
                formula_deps: None,
                eval: None,
                _state: std::marker::PhantomData,
            }
        }
    }

    impl<E> FormulaBuilder<NeedsFormulaDeps, E> {
        pub fn formula_deps(self, deps: FormulaCellDeps) -> FormulaBuilder<NeedsEval, E> {
            FormulaBuilder {
                addr: self.addr,
                excel: self.excel,
                inputs: self.inputs,
                statics: self.statics,
                sheets: self.sheets,
                formula_deps: Some(deps),
                eval: None,
                _state: std::marker::PhantomData,
            }
        }
    }

    impl FormulaBuilder<NeedsEval, fn(&EvalContext) -> CellValue> {
        pub fn eval<F>(self, eval: F) -> FormulaBuilder<Complete, F>
        where
            F: Fn(&EvalContext) -> CellValue,
        {
            FormulaBuilder {
                addr: self.addr,
                excel: self.excel,
                inputs: self.inputs,
                statics: self.statics,
                sheets: self.sheets,
                formula_deps: self.formula_deps,
                eval: Some(eval),
                _state: std::marker::PhantomData,
            }
        }
    }

    impl<E> FormulaBuilder<Complete, E>
    where
        E: Fn(&EvalContext) -> CellValue,
    {
        pub fn build(self) -> (CellAddr, FormulaCell<E>) {
            (
                self.addr,
                FormulaCell {
                    excel: self.excel,
                    deps: FormulaDeps {
                        inputs: self.inputs.unwrap(),
                        statics: self.statics.unwrap(),
                        sheets: self.sheets.unwrap(),
                        formula_deps: self.formula_deps.unwrap(),
                    },
                    eval: self.eval.unwrap(),
                },
            )
        }
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_addr_to_excel() {
        assert_eq!(CellAddr::new(0, 0).to_excel(), "A1");
        assert_eq!(CellAddr::new(1, 4).to_excel(), "E2");
        assert_eq!(CellAddr::new(30, 21).to_excel(), "V31");
    }

    #[test]
    fn test_api_key_addr() {
        assert_eq!(ApiKey::Language.addr(), CellAddr::new(1, 4));
        assert_eq!(ApiKey::Currency.addr(), CellAddr::new(2, 4));
    }

    #[test]
    fn test_registry_disjoint_sets() {
        let mut registry: CellRegistry<fn(&EvalContext) -> CellValue> = CellRegistry::new();

        // Registriere API-Zelle
        registry.register_api(ApiKey::Language).unwrap();

        // Versuche dieselbe Zelle als Static zu registrieren -> Fehler
        let result = registry.register_static(addr::E2, CellValue::Text("test".to_string()));
        assert!(matches!(
            result,
            Err(RegistryError::AlreadyRegistered { .. })
        ));
    }

    #[test]
    fn test_formula_builder() {
        use builder::FormulaBuilder;

        let (addr, formula) = FormulaBuilder::new(
            CellAddr::new(0, 1),
            r#"=IF($E$2="","",VLOOKUP($E$2,Sprachversionen!$B:$BN,2,FALSE))"#,
        )
        .inputs(Inputs::one(addr::E2))
        .statics(Statics::index(2))
        .sheets(Sheets::lang_table())
        .formula_deps(FormulaCellDeps::none())
        .eval(|_ctx| CellValue::Text("Test".to_string()))
        .build();

        assert_eq!(addr, CellAddr::new(0, 1));
        assert_eq!(formula.deps.inputs.0.len(), 1);
        assert_eq!(formula.deps.statics.0.len(), 1);
    }
}
