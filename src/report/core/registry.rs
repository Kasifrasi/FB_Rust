//! Cell Registry - Zentrale Verwaltung aller Zellen
//!
//! Verwaltet alle Zellen des Berichts (statischer Bereich, Pre-Body, Body, Footer).
//! Jede registrierte Zelle gehört zu genau einer Kategorie:
//! - Api: Wert kommt von außen (ReportValues)
//! - Static: Fester Wert im Code
//! - Formula: Hat Excel-Formel + Rust-Evaluator
//! - Empty: Nicht registriert = Standardwert
//!
//! Invarianten:
//! - Api ∩ Static = ∅
//! - Api ∩ Formula = ∅
//! - Static ∩ Formula = ∅
//! - Formel-Dependencies bilden DAG (kein Zyklus)

use crate::report::api::{CellValue, ReportValues};
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
    pub fn to_excel(self) -> String {
        let col_letter = Self::col_to_letter(self.col);
        format!("{}{}", col_letter, self.row + 1)
    }

    pub(crate) fn col_to_letter(col: u16) -> String {
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
}

// ============================================================================
// API Keys - Re-export aus api.rs (EINZIGE QUELLE DER WAHRHEIT)
// ============================================================================

// ApiKey wird jetzt in api.rs definiert und von dort exportiert.
// Siehe api.rs für die vollständige Definition aller API-Zellen.
pub use crate::report::api::ApiKey;

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

    /// Registriert eine API-Zelle (für statische Keys mit fester Adresse)
    ///
    /// Für dynamische Keys (wie `Position`) verwende `register_api_at()`.
    pub fn register_api(&mut self, key: ApiKey) -> Result<(), RegistryError> {
        let addr = key.addr();
        self.check_not_registered(addr, "Api")?;

        self.cells.insert(addr, CellKind::Api(ApiCell { key }));
        self.api_cells.insert(addr);
        self.eval_order = None; // Cache invalidieren

        Ok(())
    }

    /// Registriert eine API-Zelle an einer spezifischen Adresse
    ///
    /// Wird für dynamische Keys verwendet, deren Adresse erst zur Laufzeit
    /// bekannt ist (z.B. `ApiKey::Position` basierend auf `BodyLayout`).
    ///
    /// # Arguments
    /// * `key` - Der API-Key (kann statisch oder dynamisch sein)
    /// * `addr` - Die Zelladresse (bei dynamischen Keys aus Layout berechnet)
    pub fn register_api_at(&mut self, key: ApiKey, addr: CellAddr) -> Result<(), RegistryError> {
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

    /// Gibt die Zelle zurück (oder None wenn nicht registriert)
    pub fn get(&self, addr: CellAddr) -> Option<&CellKind<E>> {
        self.cells.get(&addr)
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

    /// Berechnet die topologische Sortierung der Formeln (mit Cache)
    pub fn compute_eval_order(&mut self) -> Result<&[CellAddr], RegistryError> {
        if self.eval_order.is_some() {
            return Ok(self.eval_order.as_ref().unwrap());
        }

        let order = self.topological_sort()?;
        self.eval_order = Some(order);
        Ok(self.eval_order.as_ref().unwrap())
    }

    /// Gibt die topologische Sortierung zurück (ohne Cache zu ändern)
    pub fn get_eval_order(&self) -> Result<Vec<CellAddr>, RegistryError> {
        if let Some(order) = &self.eval_order {
            Ok(order.clone())
        } else {
            self.topological_sort()
        }
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
            if let Some(CellKind::Formula(f)) = self.cells.get(addr) {
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

            if let Some(neighbors) = graph.get(&addr) {
                for &neighbor in neighbors {
                    let deg = in_degree.get_mut(&neighbor).unwrap();
                    *deg -= 1;
                    if *deg == 0 {
                        queue.push(neighbor);
                    }
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

    // ========================================================================
    // Dynamische Registrierung (für Body, Prebody, Footer)
    // ========================================================================

    /// Registriert eine dynamische Formel-Zelle
    ///
    /// Anders als `register_formula()` validiert diese Methode nicht, dass
    /// alle Formel-Dependencies bereits registriert sind. Dies ermöglicht
    /// die Registrierung von Formeln deren Abhängigkeiten erst später
    /// (z.B. in derselben Batch-Registrierung) hinzugefügt werden.
    ///
    /// **Verwende diese Methode für:**
    /// - Body-Formeln (Ratio, SUM)
    /// - Prebody-Formeln (VLOOKUP)
    /// - Footer-Formeln (Check, Diff)
    ///
    /// # Arguments
    /// * `addr` - Die Zelladresse
    /// * `formula` - Die Formel-Zelle
    ///
    /// # Errors
    /// - `AlreadyRegistered` wenn die Zelle bereits registriert ist
    pub fn register_dynamic_formula(
        &mut self,
        addr: CellAddr,
        formula: FormulaCell<E>,
    ) -> Result<(), RegistryError> {
        self.check_not_registered(addr, "DynamicFormula")?;

        self.cells.insert(addr, CellKind::Formula(formula));
        self.formula_cells.insert(addr);
        self.eval_order = None; // Cache invalidieren

        Ok(())
    }

    /// Registriert mehrere dynamische API-Zellen auf einmal
    ///
    /// Effizienter als mehrere einzelne `register_api_at()` Aufrufe,
    /// da der Cache nur einmal invalidiert wird.
    ///
    /// # Arguments
    /// * `cells` - Iterator über (ApiKey, CellAddr) Paare
    pub fn register_dynamic_api_batch(
        &mut self,
        cells: impl IntoIterator<Item = (ApiKey, CellAddr)>,
    ) -> Result<(), RegistryError> {
        for (key, addr) in cells {
            self.check_not_registered(addr, "Api")?;
            self.cells.insert(addr, CellKind::Api(ApiCell { key }));
            self.api_cells.insert(addr);
        }
        self.eval_order = None; // Cache nur einmal invalidieren
        Ok(())
    }

    /// Gibt die Anzahl der dynamisch registrierten Zellen zurück
    ///
    /// Dynamische Zellen sind alle Zellen mit Adressen >= Zeile 26
    /// (der erste mögliche Body-Bereich nach dem statischen A1:V25).
    pub fn dynamic_cell_count(&self) -> usize {
        self.cells.keys().filter(|addr| addr.row >= 26).count()
    }

    /// Prüft ob die Registry dynamische Zellen enthält
    pub fn has_dynamic_cells(&self) -> bool {
        self.cells.keys().any(|addr| addr.row >= 26)
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
///
/// Stellt einheitliche Zugriffsmethoden für alle Formeln bereit,
/// sowohl für statische als auch für dynamische Bereiche.
pub struct EvalContext<'a> {
    /// Bereits berechnete Werte
    pub computed: &'a HashMap<CellAddr, CellValue>,
    /// API-Werte
    pub api_values: &'a ReportValues,
}

impl<'a> EvalContext<'a> {
    /// Erstellt einen neuen Evaluierungskontext
    pub fn new(computed: &'a HashMap<CellAddr, CellValue>, api_values: &'a ReportValues) -> Self {
        Self {
            computed,
            api_values,
        }
    }

    /// Liest den Wert einer Zelle
    pub fn cell(&self, addr: CellAddr) -> &CellValue {
        static EMPTY: CellValue = CellValue::Empty;
        self.computed.get(&addr).unwrap_or(&EMPTY)
    }

    /// Liest einen API-Wert direkt
    pub fn api(&self) -> &ReportValues {
        self.api_values
    }

    // ========================================================================
    // Einheitliche Hilfsmethoden für alle Bereiche (statisch + dynamisch)
    // ========================================================================

    /// Gibt die aktuell ausgewählte Sprache zurück
    ///
    /// Zentrale Methode für alle VLOOKUP-abhängigen Formeln.
    pub fn language(&self) -> Option<&str> {
        self.api_values.language()
    }

    /// Evaluiert einen VLOOKUP-Index für die aktuelle Sprache
    ///
    /// Einheitliche Methode für alle Bereiche des Reports.
    /// Ersetzt direkte `lookup_text_string()` Aufrufe.
    ///
    /// # Arguments
    /// * `index` - Der VLOOKUP-Index (1-basiert, wie in Excel)
    /// * `default` - Optionaler Default-Wert falls nicht gefunden
    ///
    /// # Returns
    /// Der Text aus TEXT_MATRIX, Default-Wert, oder Empty
    pub fn vlookup_text(&self, index: usize) -> CellValue {
        super::definitions::lookup_text(self.language(), index)
    }

    /// Evaluiert einen VLOOKUP-Index und gibt den Text als String zurück
    ///
    /// Convenience-Methode für Fälle wo nur der String-Wert benötigt wird.
    ///
    /// # Arguments
    /// * `index` - Der VLOOKUP-Index (1-basiert)
    /// * `default` - Optionaler Default-String falls nicht gefunden
    pub fn vlookup_text_string(&self, index: usize, default: Option<&str>) -> String {
        match self.vlookup_text(index) {
            CellValue::Text(s) => s,
            CellValue::Empty => default.unwrap_or("").to_string(),
            _ => default.unwrap_or("").to_string(),
        }
    }

    /// Berechnet IFERROR(numerator/denominator, 0)
    ///
    /// Einheitliche Methode für Ratio-Berechnungen in allen Bereichen.
    /// Unterstützt sowohl Zell-Adressen als auch direkte Werte.
    pub fn iferror_division(&self, numerator: CellAddr, denominator: CellAddr) -> CellValue {
        let num = self.cell(numerator).as_number();
        let denom = self.cell(denominator).as_number();

        CellValue::Number(Self::safe_divide(num.unwrap_or(0.0), denom.unwrap_or(0.0)))
    }

    /// Sichere Division mit Null-Check
    ///
    /// Gibt 0.0 zurück wenn Denominator 0 ist, sonst numerator/denominator.
    /// Kann als Hilfsfunktion für alle Division-Operationen genutzt werden.
    #[inline]
    pub fn safe_divide(numerator: f64, denominator: f64) -> f64 {
        if denominator == 0.0 {
            0.0
        } else {
            numerator / denominator
        }
    }

    /// Berechnet SUM(range)
    ///
    /// Summiert alle Werte im Bereich.
    pub fn sum_range(&self, start: CellAddr, end: CellAddr) -> CellValue {
        let mut sum = 0.0;
        for row in start.row..=end.row {
            let addr = CellAddr::new(row, start.col);
            if let Some(n) = self.cell(addr).as_number() {
                sum += n;
            }
        }
        CellValue::Number(sum)
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
// Convenience type aliases
// ============================================================================

/// Boxed dynamic evaluator closure for use with [`CellRegistry`].
pub type DynEvaluator = Box<dyn Fn(&EvalContext) -> CellValue>;

/// [`CellRegistry`] with a boxed dynamic evaluator — the concrete type used throughout the codebase.
pub type DynRegistry = CellRegistry<DynEvaluator>;

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
