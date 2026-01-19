//! Zell-Definitionen für den Finanzbericht
//!
//! Dieses Modul definiert alle Zellen des Finanzberichts typsicher.
//! Jede Zelle hat eine eindeutige Adresse (Row, Col) und einen Typ.

use std::fmt;

/// Excel-Zelladresse (0-basiert)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellAddress {
    pub row: u32,
    pub col: u16,
}

impl CellAddress {
    pub const fn new(row: u32, col: u16) -> Self {
        Self { row, col }
    }

    /// Konvertiert zu Excel-Notation (z.B. "E2")
    pub fn to_excel_notation(&self) -> String {
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

impl fmt::Display for CellAddress {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_excel_notation())
    }
}

// ============================================================================
// Header Input Cells - Zellen die Benutzereingaben akzeptieren
// ============================================================================

/// Alle Eingabezellen im Header-Bereich
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HeaderInputCell {
    /// E2: Sprachauswahl (Dropdown)
    Language,
    /// E3: Währungsauswahl (Dropdown)
    Currency,
    /// D5: Projektnummer (freie Eingabe)
    ProjectNumber,
    /// D6:H7: Projekttitel (merged, freie Eingabe)
    ProjectTitle,
    /// E8: Projektstart-Datum
    ProjectStartDate,
    /// G8:H8: Projektende-Datum (merged)
    ProjectEndDate,
    /// E9: Berichtszeitraum Start
    ReportPeriodStart,
    /// G9:H9: Berichtszeitraum Ende (merged)
    ReportPeriodEnd,
    /// J7: Wechselkurs-Datum (orange)
    ExchangeRateDate,
    /// J8: Wechselkurs-Wert (grau)
    ExchangeRateValue,
    /// J9: Wechselkurs-Eingabe (gelb)
    ExchangeRateInput,
}

impl HeaderInputCell {
    /// Gibt die Zelladresse zurück
    pub const fn address(&self) -> CellAddress {
        match self {
            Self::Language => CellAddress::new(1, 4),          // E2
            Self::Currency => CellAddress::new(2, 4),          // E3
            Self::ProjectNumber => CellAddress::new(4, 3),     // D5
            Self::ProjectTitle => CellAddress::new(5, 3),      // D6
            Self::ProjectStartDate => CellAddress::new(7, 4),  // E8
            Self::ProjectEndDate => CellAddress::new(7, 6),    // G8
            Self::ReportPeriodStart => CellAddress::new(8, 4), // E9
            Self::ReportPeriodEnd => CellAddress::new(8, 6),   // G9
            Self::ExchangeRateDate => CellAddress::new(6, 9),  // J7
            Self::ExchangeRateValue => CellAddress::new(7, 9), // J8
            Self::ExchangeRateInput => CellAddress::new(8, 9), // J9
        }
    }

    /// Gibt alle Header-Eingabezellen zurück
    pub fn all() -> &'static [HeaderInputCell] {
        &[
            Self::Language,
            Self::Currency,
            Self::ProjectNumber,
            Self::ProjectTitle,
            Self::ProjectStartDate,
            Self::ProjectEndDate,
            Self::ReportPeriodStart,
            Self::ReportPeriodEnd,
            Self::ExchangeRateDate,
            Self::ExchangeRateValue,
            Self::ExchangeRateInput,
        ]
    }
}

// ============================================================================
// Table Input Cells - Eingabezellen im Tabellen-Body
// ============================================================================

/// Eingabezellen im linken Tabellen-Body (Einnahmen/Ausgaben)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TableInputCell {
    /// D15-D19: Bewilligtes Budget (Zeile 0-4)
    ApprovedBudget(u8),
    /// E15-E19: Einnahmen im Berichtszeitraum (Zeile 0-4)
    IncomeReportPeriod(u8),
    /// F15-F19: Einnahmen gesamt (Zeile 0-4)
    IncomeTotal(u8),
    /// H15-H19: Begründung für Abweichung (Zeile 0-4)
    IncomeReason(u8),
}

impl TableInputCell {
    pub const fn address(&self) -> CellAddress {
        match self {
            Self::ApprovedBudget(row) => CellAddress::new(14 + *row as u32, 3), // D15+
            Self::IncomeReportPeriod(row) => CellAddress::new(14 + *row as u32, 4), // E15+
            Self::IncomeTotal(row) => CellAddress::new(14 + *row as u32, 5),    // F15+
            Self::IncomeReason(row) => CellAddress::new(14 + *row as u32, 7),   // H15+
        }
    }
}

// ============================================================================
// Right Panel Input Cells - Eingabezellen im rechten Panel
// ============================================================================

/// Eingabezellen im Right Panel (18 Zeilen pro Seite, 2 Seiten)
///
/// NUR diese Zellen sind API-Eingabefelder:
/// - L14:N31 (Datum, Euro, Lokal - linke Seite)
/// - S14:U31 (Datum, Euro, Lokal - rechte Seite)
///
/// NICHT API (sind Formeln):
/// - K14:K31 (Nummern links - Formeln!)
/// - R14:R31 (Nummern rechts - Formeln!)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RightPanelInputCell {
    /// L14-L31: Datum (linke Seite, Index 0-17)
    LeftDate(u8),
    /// M14-M31: Betrag Euro (linke Seite, Index 0-17)
    LeftAmount1(u8),
    /// N14-N31: Betrag Lokal (linke Seite, Index 0-17)
    LeftAmount2(u8),
    /// S14-S31: Datum (rechte Seite, Index 0-17)
    RightDate(u8),
    /// T14-T31: Betrag Euro (rechte Seite, Index 0-17)
    RightAmount1(u8),
    /// U14-U31: Betrag Lokal (rechte Seite, Index 0-17)
    RightAmount2(u8),
}

impl RightPanelInputCell {
    pub const fn address(&self) -> CellAddress {
        match self {
            Self::LeftDate(idx) => CellAddress::new(13 + *idx as u32, 11), // L14+
            Self::LeftAmount1(idx) => CellAddress::new(13 + *idx as u32, 12), // M14+
            Self::LeftAmount2(idx) => CellAddress::new(13 + *idx as u32, 13), // N14+
            Self::RightDate(idx) => CellAddress::new(13 + *idx as u32, 18), // S14+
            Self::RightAmount1(idx) => CellAddress::new(13 + *idx as u32, 19), // T14+
            Self::RightAmount2(idx) => CellAddress::new(13 + *idx as u32, 20), // U14+
        }
    }
}

// ============================================================================
// Unified Input Cell Enum
// ============================================================================

/// Alle Eingabezellen des Finanzberichts
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum InputCell {
    Header(HeaderInputCell),
    Table(TableInputCell),
    RightPanel(RightPanelInputCell),
}

impl InputCell {
    pub fn address(&self) -> CellAddress {
        match self {
            Self::Header(cell) => cell.address(),
            Self::Table(cell) => cell.address(),
            Self::RightPanel(cell) => cell.address(),
        }
    }
}

impl From<HeaderInputCell> for InputCell {
    fn from(cell: HeaderInputCell) -> Self {
        Self::Header(cell)
    }
}

impl From<TableInputCell> for InputCell {
    fn from(cell: TableInputCell) -> Self {
        Self::Table(cell)
    }
}

impl From<RightPanelInputCell> for InputCell {
    fn from(cell: RightPanelInputCell) -> Self {
        Self::RightPanel(cell)
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_address_to_excel() {
        assert_eq!(CellAddress::new(0, 0).to_excel_notation(), "A1");
        assert_eq!(CellAddress::new(1, 4).to_excel_notation(), "E2");
        assert_eq!(CellAddress::new(13, 11).to_excel_notation(), "L14");
        assert_eq!(CellAddress::new(0, 25).to_excel_notation(), "Z1");
        assert_eq!(CellAddress::new(0, 26).to_excel_notation(), "AA1");
    }

    #[test]
    fn test_header_input_cell_addresses() {
        assert_eq!(HeaderInputCell::Language.address(), CellAddress::new(1, 4));
        assert_eq!(HeaderInputCell::Currency.address(), CellAddress::new(2, 4));
    }
}
