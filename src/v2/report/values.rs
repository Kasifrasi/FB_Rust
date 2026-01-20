//! Werte-Speicher für den Finanzbericht
//!
//! Dieses Modul speichert alle Eingabewerte.
//! Nutzt direkt `ApiKey` als Schlüssel - keine redundanten Enums.

use super::api::{ApiKey, PositionField};
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
            CellValue::Date(s) => Some(s), // Date ist auch ein String
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
///
/// Nutzt direkt `ApiKey` als Schlüssel für type-safe Zugriff.
///
/// # Beispiel
/// ```ignore
/// let mut values = ReportValues::new();
/// values.set(ApiKey::Language, "deutsch");
/// values.set(ApiKey::ApprovedBudget(0), 1000.0);
///
/// assert_eq!(values.get(ApiKey::Language).as_text(), Some("deutsch"));
/// ```
#[derive(Debug, Clone, Default)]
pub struct ReportValues {
    /// Speicher für alle Zellwerte, indexiert durch ApiKey
    values: HashMap<ApiKey, CellValue>,
}

impl ReportValues {
    pub fn new() -> Self {
        Self::default()
    }

    /// Setzt einen Wert für eine API-Zelle
    pub fn set(&mut self, key: ApiKey, value: impl Into<CellValue>) -> &mut Self {
        self.values.insert(key, value.into());
        self
    }

    /// Holt einen Wert für eine API-Zelle
    pub fn get(&self, key: ApiKey) -> &CellValue {
        self.values.get(&key).unwrap_or(&CellValue::Empty)
    }

    /// Prüft ob eine Zelle einen Wert hat
    pub fn has_value(&self, key: ApiKey) -> bool {
        self.values
            .get(&key)
            .map(|v| !v.is_empty())
            .unwrap_or(false)
    }

    // ========================================================================
    // Convenience Getter für häufig verwendete Zellen
    // ========================================================================

    /// Gibt die ausgewählte Sprache zurück (E2)
    pub fn language(&self) -> Option<&str> {
        self.get(ApiKey::Language).as_text()
    }

    /// Gibt die ausgewählte Währung zurück (E3)
    pub fn currency(&self) -> Option<&str> {
        self.get(ApiKey::Currency).as_text()
    }

    /// Gibt die Projektnummer zurück (D5)
    pub fn project_number(&self) -> Option<&str> {
        self.get(ApiKey::ProjectNumber).as_text()
    }

    /// Gibt den Projekttitel zurück (D6)
    pub fn project_title(&self) -> Option<&str> {
        self.get(ApiKey::ProjectTitle).as_text()
    }

    // ========================================================================
    // Convenience Setter (Builder-Pattern)
    // ========================================================================

    /// Setzt die Sprache (E2)
    pub fn with_language(mut self, lang: &str) -> Self {
        self.set(ApiKey::Language, lang);
        self
    }

    /// Setzt die Währung (E3)
    pub fn with_currency(mut self, currency: &str) -> Self {
        self.set(ApiKey::Currency, currency);
        self
    }

    /// Setzt die Projektnummer (D5)
    pub fn with_project_number(mut self, number: &str) -> Self {
        self.set(ApiKey::ProjectNumber, number);
        self
    }

    /// Setzt den Projekttitel (D6)
    pub fn with_project_title(mut self, title: &str) -> Self {
        self.set(ApiKey::ProjectTitle, title);
        self
    }

    /// Setzt Projektstart-Datum (E8)
    pub fn with_project_start(mut self, date: &str) -> Self {
        self.set(ApiKey::ProjectStart, CellValue::Date(date.to_string()));
        self
    }

    /// Setzt Projektende-Datum (G8)
    pub fn with_project_end(mut self, date: &str) -> Self {
        self.set(ApiKey::ProjectEnd, CellValue::Date(date.to_string()));
        self
    }

    /// Setzt Berichtszeitraum Start (E9)
    pub fn with_report_start(mut self, date: &str) -> Self {
        self.set(ApiKey::ReportStart, CellValue::Date(date.to_string()));
        self
    }

    /// Setzt Berichtszeitraum Ende (G9)
    pub fn with_report_end(mut self, date: &str) -> Self {
        self.set(ApiKey::ReportEnd, CellValue::Date(date.to_string()));
        self
    }

    // ========================================================================
    // Positions-Methoden (Dynamische Kostenpositionen)
    // ========================================================================

    /// Setzt ein einzelnes Positions-Feld
    ///
    /// # Arguments
    /// * `category` - Kategorie-Nummer (1-8)
    /// * `position` - Position:
    ///   - `0`: Header-Eingabe (bei Kategorien mit 0 Positionen in BodyConfig)
    ///   - `1..N`: Positions-Zeile (bei Kategorien mit 1+ Positionen)
    /// * `field` - Welches Feld der Position
    /// * `value` - Der Wert
    ///
    /// # Hinweis
    /// Bei `position=0` (Header-Eingabe) ist `PositionField::Description` nicht
    /// verfügbar, da Spalte C das VLOOKUP-Label enthält.
    ///
    /// # Beispiel
    /// ```ignore
    /// // Position mit Positions-Zeile (position >= 1)
    /// values.set_position(1, 1, PositionField::Description, "Personalkosten");
    /// values.set_position(1, 1, PositionField::Approved, 5000.0);
    ///
    /// // Header-Eingabe (position = 0)
    /// values.set_position(6, 0, PositionField::Approved, 3000.0);
    /// ```
    pub fn set_position(
        &mut self,
        category: u8,
        position: u16,
        field: PositionField,
        value: impl Into<CellValue>,
    ) -> &mut Self {
        self.set(
            ApiKey::Position {
                category,
                position,
                field,
            },
            value,
        )
    }

    /// Setzt eine komplette Positions-Zeile (alle 5 Felder)
    ///
    /// **Nur für position >= 1!** Für Header-Eingabe (position=0) verwende
    /// `set_header_input()`.
    ///
    /// # Arguments
    /// * `category` - Kategorie-Nummer (1-8)
    /// * `position` - Position innerhalb der Kategorie (1-basiert!)
    /// * `description` - Beschreibung (Spalte C)
    /// * `approved` - Bewilligtes Budget (Spalte D)
    /// * `income_report` - Einnahmen Berichtszeitraum (Spalte E)
    /// * `income_total` - Einnahmen gesamt (Spalte F)
    /// * `remark` - Begründung/Bemerkung (Spalte H)
    ///
    /// # Beispiel
    /// ```ignore
    /// values.set_position_row(1, 1, "Personalkosten", 5000.0, 2500.0, 2500.0, "");
    /// ```
    pub fn set_position_row(
        &mut self,
        category: u8,
        position: u16,
        description: impl Into<CellValue>,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        use PositionField::*;
        self.set_position(category, position, Description, description);
        self.set_position(category, position, Approved, approved);
        self.set_position(category, position, IncomeReport, income_report);
        self.set_position(category, position, IncomeTotal, income_total);
        self.set_position(category, position, Remark, remark);
        self
    }

    /// Setzt Header-Eingabe-Werte (position=0)
    ///
    /// Für Kategorien mit 0 Positionen in BodyConfig.
    /// Kein Description-Feld (C ist VLOOKUP-Label).
    ///
    /// # Arguments
    /// * `category` - Kategorie-Nummer (1-8)
    /// * `approved` - Bewilligtes Budget (Spalte D)
    /// * `income_report` - Einnahmen Berichtszeitraum (Spalte E)
    /// * `income_total` - Einnahmen gesamt (Spalte F)
    /// * `remark` - Begründung/Bemerkung (Spalte H)
    ///
    /// # Beispiel
    /// ```ignore
    /// // Kategorie 6 mit 0 Positionen in BodyConfig
    /// values.set_header_input(6, 4000.0, 2000.0, 2000.0, "Sonstiges");
    /// ```
    pub fn set_header_input(
        &mut self,
        category: u8,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        use PositionField::*;
        // position=0 für Header-Eingabe
        self.set_position(category, 0, Approved, approved);
        self.set_position(category, 0, IncomeReport, income_report);
        self.set_position(category, 0, IncomeTotal, income_total);
        self.set_position(category, 0, Remark, remark);
        self
    }

    /// Holt einen Positions-Wert (falls vorhanden)
    pub fn get_position(&self, category: u8, position: u16, field: PositionField) -> &CellValue {
        self.get(ApiKey::Position {
            category,
            position,
            field,
        })
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
        values.set(ApiKey::Language, "english");
        values.set(ApiKey::ApprovedBudget(0), 1000.0);

        assert_eq!(values.get(ApiKey::Language).as_text(), Some("english"));
        assert_eq!(
            values.get(ApiKey::ApprovedBudget(0)).as_number(),
            Some(1000.0)
        );
    }

    #[test]
    fn test_range_cells() {
        let mut values = ReportValues::new();

        // Setze alle 5 ApprovedBudget Werte
        for i in 0..5u8 {
            values.set(ApiKey::ApprovedBudget(i), 1000.0 * (i + 1) as f64);
        }

        assert_eq!(
            values.get(ApiKey::ApprovedBudget(0)).as_number(),
            Some(1000.0)
        );
        assert_eq!(
            values.get(ApiKey::ApprovedBudget(4)).as_number(),
            Some(5000.0)
        );
    }

    #[test]
    fn test_right_panel_cells() {
        let mut values = ReportValues::new();

        values.set(
            ApiKey::LeftDate(0),
            CellValue::Date("2024-01-15".to_string()),
        );
        values.set(ApiKey::LeftAmountEuro(0), 500.0);
        values.set(ApiKey::RightAmountLocal(17), 999.99);

        assert_eq!(
            values.get(ApiKey::LeftDate(0)).as_text(),
            Some("2024-01-15")
        );
        assert_eq!(
            values.get(ApiKey::LeftAmountEuro(0)).as_number(),
            Some(500.0)
        );
        assert_eq!(
            values.get(ApiKey::RightAmountLocal(17)).as_number(),
            Some(999.99)
        );
    }

    #[test]
    fn test_set_position() {
        use PositionField::*;

        let mut values = ReportValues::new();
        values.set_position(1, 1, Description, "Personalkosten");
        values.set_position(1, 1, Approved, 5000.0);

        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Personalkosten")
        );
        assert_eq!(
            values.get_position(1, 1, Approved).as_number(),
            Some(5000.0)
        );
    }

    #[test]
    fn test_set_position_row() {
        use PositionField::*;

        let mut values = ReportValues::new();
        values.set_position_row(
            1,                    // Kategorie
            1,                    // Position
            "Reisekosten",        // Description
            2000.0,               // Approved
            1800.0,               // IncomeReport
            1800.0,               // IncomeTotal
            "Dienstreise Berlin", // Remark
        );

        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Reisekosten")
        );
        assert_eq!(
            values.get_position(1, 1, Approved).as_number(),
            Some(2000.0)
        );
        assert_eq!(
            values.get_position(1, 1, IncomeReport).as_number(),
            Some(1800.0)
        );
        assert_eq!(
            values.get_position(1, 1, IncomeTotal).as_number(),
            Some(1800.0)
        );
        assert_eq!(
            values.get_position(1, 1, Remark).as_text(),
            Some("Dienstreise Berlin")
        );
    }

    #[test]
    fn test_header_input() {
        use PositionField::*;

        let mut values = ReportValues::new();
        values.set_header_input(6, 4000.0, 2000.0, 2000.0, "Sonstiges");

        // position=0 für Header-Eingabe
        assert_eq!(
            values.get_position(6, 0, Approved).as_number(),
            Some(4000.0)
        );
        assert_eq!(
            values.get_position(6, 0, IncomeReport).as_number(),
            Some(2000.0)
        );
        assert_eq!(
            values.get_position(6, 0, IncomeTotal).as_number(),
            Some(2000.0)
        );
        assert_eq!(
            values.get_position(6, 0, Remark).as_text(),
            Some("Sonstiges")
        );

        // Description bei Header-Eingabe nicht gesetzt
        assert!(values.get_position(6, 0, Description).is_empty());
    }

    #[test]
    fn test_multiple_positions() {
        use PositionField::*;

        let mut values = ReportValues::new();

        // Kategorie 1, Positionen 1-3
        values.set_position_row(1, 1, "Personal", 5000.0, 2500.0, 2500.0, "");
        values.set_position_row(1, 2, "Reisen", 2000.0, 1800.0, 1800.0, "");
        values.set_position_row(1, 3, "Material", 1000.0, 500.0, 500.0, "");

        // Kategorie 2, Position 1
        values.set_position_row(2, 1, "Externe", 3000.0, 1500.0, 1500.0, "");

        // Prüfe verschiedene Positionen
        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Personal")
        );
        assert_eq!(
            values.get_position(1, 3, Approved).as_number(),
            Some(1000.0)
        );
        assert_eq!(
            values.get_position(2, 1, Description).as_text(),
            Some("Externe")
        );

        // Nicht gesetzte Position gibt Empty zurück
        assert!(values.get_position(1, 4, Description).is_empty());
    }

    #[test]
    fn test_mixed_modes() {
        use PositionField::*;

        let mut values = ReportValues::new();

        // Kategorie 1: Positions-Modus (position >= 1)
        values.set_position_row(1, 1, "Personal", 5000.0, 2500.0, 2500.0, "");

        // Kategorie 6: Header-Eingabe-Modus (position = 0)
        values.set_header_input(6, 4000.0, 2000.0, 2000.0, "");

        // Prüfe beide Modi
        assert_eq!(
            values.get_position(1, 1, Description).as_text(),
            Some("Personal")
        );
        assert_eq!(
            values.get_position(6, 0, Approved).as_number(),
            Some(4000.0)
        );
    }
}
