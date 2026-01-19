//! Werte-Speicher für den Finanzbericht
//!
//! Dieses Modul speichert alle Eingabewerte und berechneten Werte.

use super::cells::{HeaderInputCell, InputCell};
use std::collections::HashMap;

// ============================================================================
// Cell Value Types
// ============================================================================

/// Mögliche Werte einer Zelle
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    Date(String), // Format: "YYYY-MM-DD" oder "DD.MM.YYYY"
}

impl CellValue {
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    pub fn as_text(&self) -> Option<&str> {
        match self {
            CellValue::Text(s) => Some(s),
            _ => None,
        }
    }

    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            _ => None,
        }
    }
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        if s.is_empty() {
            CellValue::Empty
        } else {
            CellValue::Text(s.to_string())
        }
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        if s.is_empty() {
            CellValue::Empty
        } else {
            CellValue::Text(s)
        }
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

// ============================================================================
// Report Values - Alle Eingabewerte des Finanzberichts
// ============================================================================

/// Speichert alle Eingabewerte eines Finanzberichts
#[derive(Debug, Clone, Default)]
pub struct ReportValues {
    /// Generischer Speicher für alle Zellwerte
    values: HashMap<InputCell, CellValue>,
}

impl ReportValues {
    pub fn new() -> Self {
        Self::default()
    }

    /// Setzt einen Wert für eine Eingabezelle
    pub fn set<C: Into<InputCell>, V: Into<CellValue>>(&mut self, cell: C, value: V) -> &mut Self {
        self.values.insert(cell.into(), value.into());
        self
    }

    /// Holt einen Wert für eine Eingabezelle
    pub fn get<C: Into<InputCell>>(&self, cell: C) -> &CellValue {
        self.values.get(&cell.into()).unwrap_or(&CellValue::Empty)
    }

    /// Prüft ob eine Zelle einen Wert hat
    pub fn has_value<C: Into<InputCell>>(&self, cell: C) -> bool {
        self.values
            .get(&cell.into())
            .map(|v| !v.is_empty())
            .unwrap_or(false)
    }

    // ========================================================================
    // Convenience Getter für Header-Zellen
    // ========================================================================

    /// Gibt die ausgewählte Sprache zurück (E2)
    pub fn language(&self) -> Option<&str> {
        self.get(HeaderInputCell::Language).as_text()
    }

    /// Gibt die ausgewählte Währung zurück (E3)
    pub fn currency(&self) -> Option<&str> {
        self.get(HeaderInputCell::Currency).as_text()
    }

    /// Gibt die Projektnummer zurück (D5)
    pub fn project_number(&self) -> Option<&str> {
        self.get(HeaderInputCell::ProjectNumber).as_text()
    }

    /// Gibt den Projekttitel zurück (D6)
    pub fn project_title(&self) -> Option<&str> {
        self.get(HeaderInputCell::ProjectTitle).as_text()
    }

    // ========================================================================
    // Convenience Setter (Builder-Pattern)
    // ========================================================================

    /// Setzt die Sprache (E2)
    pub fn with_language(mut self, lang: &str) -> Self {
        self.set(HeaderInputCell::Language, lang);
        self
    }

    /// Setzt die Währung (E3)
    pub fn with_currency(mut self, currency: &str) -> Self {
        self.set(HeaderInputCell::Currency, currency);
        self
    }

    /// Setzt die Projektnummer (D5)
    pub fn with_project_number(mut self, number: &str) -> Self {
        self.set(HeaderInputCell::ProjectNumber, number);
        self
    }

    /// Setzt den Projekttitel (D6)
    pub fn with_project_title(mut self, title: &str) -> Self {
        self.set(HeaderInputCell::ProjectTitle, title);
        self
    }

    /// Setzt Projektstart-Datum (E8)
    pub fn with_project_start(mut self, date: &str) -> Self {
        self.set(
            HeaderInputCell::ProjectStartDate,
            CellValue::Date(date.to_string()),
        );
        self
    }

    /// Setzt Projektende-Datum (G8)
    pub fn with_project_end(mut self, date: &str) -> Self {
        self.set(
            HeaderInputCell::ProjectEndDate,
            CellValue::Date(date.to_string()),
        );
        self
    }

    /// Setzt Berichtszeitraum Start (E9)
    pub fn with_report_start(mut self, date: &str) -> Self {
        self.set(
            HeaderInputCell::ReportPeriodStart,
            CellValue::Date(date.to_string()),
        );
        self
    }

    /// Setzt Berichtszeitraum Ende (G9)
    pub fn with_report_end(mut self, date: &str) -> Self {
        self.set(
            HeaderInputCell::ReportPeriodEnd,
            CellValue::Date(date.to_string()),
        );
        self
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_report_values_builder() {
        let values = ReportValues::new()
            .with_language("deutsch")
            .with_currency("EUR")
            .with_project_number("12345");

        assert_eq!(values.language(), Some("deutsch"));
        assert_eq!(values.currency(), Some("EUR"));
        assert_eq!(values.project_number(), Some("12345"));
    }

    #[test]
    fn test_generic_set_get() {
        let mut values = ReportValues::new();
        values.set(HeaderInputCell::Language, "english");

        assert_eq!(
            values.get(HeaderInputCell::Language).as_text(),
            Some("english")
        );
    }
}
