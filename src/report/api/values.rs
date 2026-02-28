//! Werte-Speicher für den Finanzbericht
//!
//! Dieses Modul speichert alle Eingabewerte.
//! Nutzt direkt `ApiKey` als Schlüssel - keine redundanten Enums.

use super::keys::{ApiKey, FooterField, PositionField};
use super::types::{Category, Currency, Language, ReportDate};
use std::collections::HashMap;

// ============================================================================
// Cell Value Types
// ============================================================================

/// Mögliche Werte einer Zelle
#[derive(Debug, Clone, PartialEq, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub enum CellValue {
    #[default]
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
    /// Speicher für ALLE Zellwerte, indexiert durch ApiKey
    ///
    /// Einheitlicher Zugriffspunkt für:
    /// - Header-Werte (Language, Currency, etc.)
    /// - Positions-Werte (dynamisch)
    /// - Footer-Werte (Bank, Kasse, Sonstiges)
    /// - Rechtes Panel (Währungsumrechnungen)
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

    /// Holt einen Wert für eine API-Zelle (Referenz)
    ///
    /// Einheitlicher Getter für ALLE Keys (Header, Positions, Footer, Panel).
    /// Gibt `&CellValue::Empty` zurück wenn nicht gesetzt.
    pub fn get(&self, key: ApiKey) -> &CellValue {
        self.values.get(&key).unwrap_or(&CellValue::Empty)
    }

    /// Holt einen Wert für eine API-Zelle (owned)
    ///
    /// Einheitlicher Getter für ALLE Keys (Header, Positions, Footer, Panel).
    /// Gibt `CellValue::Empty` zurück wenn nicht gesetzt.
    pub fn get_owned(&self, key: ApiKey) -> CellValue {
        self.values.get(&key).cloned().unwrap_or(CellValue::Empty)
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
    // Typsichere Setter (empfohlen)
    // ========================================================================

    /// Setzt die Sprache typsicher (E2)
    ///
    /// Verwendet das `Language` enum für Compile-Zeit-Sicherheit.
    ///
    /// # Beispiel
    /// ```ignore
    /// use kmw_fb_rust::report::types::Language;
    ///
    /// let values = ReportValues::new()
    ///     .with_lang(Language::English);
    /// ```
    pub fn with_lang(mut self, lang: Language) -> Self {
        self.set(ApiKey::Language, lang.as_str());
        self
    }

    /// Setzt die Währung typsicher (E3)
    ///
    /// Verwendet den validierten `Currency` Typ.
    ///
    /// # Beispiel
    /// ```ignore
    /// use kmw_fb_rust::report::types::Currency;
    ///
    /// let values = ReportValues::new()
    ///     .with_curr(Currency::eur());
    /// // Oder mit Validierung:
    ///     .with_curr(Currency::new("USD").unwrap());
    /// ```
    pub fn with_curr(mut self, currency: Currency) -> Self {
        self.set(ApiKey::Currency, currency.as_str());
        self
    }

    /// Setzt Projektstart-Datum typsicher (E8)
    ///
    /// Verwendet das validierte `ReportDate`.
    pub fn with_project_start_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ProjectStart, CellValue::Date(date.format_de()));
        self
    }

    /// Setzt Projektende-Datum typsicher (G8)
    pub fn with_project_end_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ProjectEnd, CellValue::Date(date.format_de()));
        self
    }

    /// Setzt Berichtszeitraum Start typsicher (E9)
    pub fn with_report_start_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ReportStart, CellValue::Date(date.format_de()));
        self
    }

    /// Setzt Berichtszeitraum Ende typsicher (G9)
    pub fn with_report_end_date(mut self, date: ReportDate) -> Self {
        self.set(ApiKey::ReportEnd, CellValue::Date(date.format_de()));
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
    #[allow(clippy::too_many_arguments)]
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

    // ========================================================================
    // Typsichere Positions-Methoden mit Category enum
    // ========================================================================

    /// Setzt ein einzelnes Positions-Feld typsicher
    ///
    /// Verwendet das `Category` enum für Compile-Zeit-Sicherheit.
    ///
    /// # Beispiel
    /// ```ignore
    /// use kmw_fb_rust::report::types::Category;
    ///
    /// values.set_cat_position(Category::Personal, 1, PositionField::Approved, 5000.0);
    /// ```
    pub fn set_cat_position(
        &mut self,
        category: Category,
        position: u16,
        field: PositionField,
        value: impl Into<CellValue>,
    ) -> &mut Self {
        self.set_position(category.index(), position, field, value)
    }

    /// Setzt eine komplette Positions-Zeile typsicher
    ///
    /// # Beispiel
    /// ```ignore
    /// use kmw_fb_rust::report::types::Category;
    ///
    /// values.set_cat_position_row(
    ///     Category::Personal, 1,
    ///     "Projektleiter", 60000.0, 30000.0, 30000.0, ""
    /// );
    /// ```
    #[allow(clippy::too_many_arguments)]
    pub fn set_cat_position_row(
        &mut self,
        category: Category,
        position: u16,
        description: impl Into<CellValue>,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        self.set_position_row(
            category.index(),
            position,
            description,
            approved,
            income_report,
            income_total,
            remark,
        )
    }

    /// Setzt Header-Eingabe typsicher
    ///
    /// # Beispiel
    /// ```ignore
    /// use kmw_fb_rust::report::types::Category;
    ///
    /// values.set_cat_header_input(
    ///     Category::Projektverwaltung,
    ///     8000.0, 4000.0, 4000.0, "Verwaltungskosten"
    /// );
    /// ```
    pub fn set_cat_header_input(
        &mut self,
        category: Category,
        approved: impl Into<CellValue>,
        income_report: impl Into<CellValue>,
        income_total: impl Into<CellValue>,
        remark: impl Into<CellValue>,
    ) -> &mut Self {
        self.set_header_input(
            category.index(),
            approved,
            income_report,
            income_total,
            remark,
        )
    }

    // ========================================================================
    // Footer-Werte (Saldenabstimmung)
    // ========================================================================
    //
    // Alle Footer-Werte werden nun in der HashMap gespeichert.
    // Diese Convenience-Methoden sind nur Wrapper um set()/get().

    /// Setzt den Bank-Saldo für die Saldenabstimmung im Footer
    pub fn with_footer_bank(mut self, value: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Bank), value);
        self
    }

    /// Setzt den Kassen-Saldo für die Saldenabstimmung im Footer
    pub fn with_footer_kasse(mut self, value: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Kasse), value);
        self
    }

    /// Setzt den Sonstiges-Saldo für die Saldenabstimmung im Footer
    /// (noch nicht eingelöste Schecks, Vorschüsse, Darlehen, etc.)
    pub fn with_footer_sonstiges(mut self, value: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Sonstiges), value);
        self
    }

    /// Setzt alle Footer-Salden auf einmal
    pub fn with_footer_salden(mut self, bank: f64, kasse: f64, sonstiges: f64) -> Self {
        self.set(ApiKey::Footer(FooterField::Bank), bank);
        self.set(ApiKey::Footer(FooterField::Kasse), kasse);
        self.set(ApiKey::Footer(FooterField::Sonstiges), sonstiges);
        self
    }

    /// Setzt den Bank-Saldo (mutierend)
    pub fn set_footer_bank(&mut self, value: f64) -> &mut Self {
        self.set(ApiKey::Footer(FooterField::Bank), value);
        self
    }

    /// Setzt den Kassen-Saldo (mutierend)
    pub fn set_footer_kasse(&mut self, value: f64) -> &mut Self {
        self.set(ApiKey::Footer(FooterField::Kasse), value);
        self
    }

    /// Setzt den Sonstiges-Saldo (mutierend)
    pub fn set_footer_sonstiges(&mut self, value: f64) -> &mut Self {
        self.set(ApiKey::Footer(FooterField::Sonstiges), value);
        self
    }

    /// Holt den Bank-Saldo
    pub fn footer_bank(&self) -> Option<f64> {
        self.get(ApiKey::Footer(FooterField::Bank)).as_number()
    }

    /// Holt den Kassen-Saldo
    pub fn footer_kasse(&self) -> Option<f64> {
        self.get(ApiKey::Footer(FooterField::Kasse)).as_number()
    }

    /// Holt den Sonstiges-Saldo
    pub fn footer_sonstiges(&self) -> Option<f64> {
        self.get(ApiKey::Footer(FooterField::Sonstiges)).as_number()
    }

    // ========================================================================
    // Validierungs-Methoden (Saldenabstimmung)
    // ========================================================================

    /// Berechnet die Summe aller Footer-Salden
    ///
    /// Wird für die Saldenabstimmung verwendet.
    pub fn footer_balance_total(&self) -> f64 {
        self.footer_bank().unwrap_or(0.0)
            + self.footer_kasse().unwrap_or(0.0)
            + self.footer_sonstiges().unwrap_or(0.0)
    }

    /// Validiert ob alle erforderlichen Footer-Werte gesetzt sind
    ///
    /// Für eine vollständige Saldenabstimmung sollten mindestens
    /// Bank und Kasse gesetzt sein.
    pub fn validate_footer_complete(&self) -> bool {
        self.footer_bank().is_some() && self.footer_kasse().is_some()
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

    #[test]
    fn test_footer_values_in_hashmap() {
        let values = ReportValues::new()
            .with_footer_bank(1500.50)
            .with_footer_kasse(250.25)
            .with_footer_sonstiges(100.0);

        // Prüfe Footer-Werte über die Convenience-Methoden
        assert_eq!(values.footer_bank(), Some(1500.50));
        assert_eq!(values.footer_kasse(), Some(250.25));
        assert_eq!(values.footer_sonstiges(), Some(100.0));

        // Prüfe dass Footer-Werte in der HashMap sind (einheitlicher Zugriff!)
        assert_eq!(
            values.get(ApiKey::Footer(FooterField::Bank)).as_number(),
            Some(1500.50)
        );
        assert_eq!(
            values.get(ApiKey::Footer(FooterField::Kasse)).as_number(),
            Some(250.25)
        );
        assert_eq!(
            values
                .get(ApiKey::Footer(FooterField::Sonstiges))
                .as_number(),
            Some(100.0)
        );

        // Prüfe Saldenabstimmungs-Logik
        assert_eq!(values.footer_balance_total(), 1850.75);
        assert!(values.validate_footer_complete());
    }

    #[test]
    fn test_footer_mutating_setters() {
        let mut values = ReportValues::new();

        values
            .set_footer_bank(2000.0)
            .set_footer_kasse(500.0)
            .set_footer_sonstiges(50.0);

        assert_eq!(values.footer_balance_total(), 2550.0);
    }

    #[test]
    fn test_footer_validation_incomplete() {
        let values = ReportValues::new().with_footer_bank(1000.0);

        // Nur Bank gesetzt, Kasse fehlt
        assert!(!values.validate_footer_complete());

        let values = values.with_footer_kasse(500.0);
        assert!(values.validate_footer_complete());
    }

    #[test]
    fn test_unified_access_all_types() {
        // Dieser Test zeigt die Eleganz der HashMap-Lösung:
        // ALLE Werte (Header, Positions, Footer) über denselben Mechanismus!

        let mut values = ReportValues::new()
            .with_language("deutsch")
            .with_footer_bank(1000.0);

        values.set_position(1, 1, PositionField::Approved, 5000.0);

        // Einheitlicher Zugriff über get()
        assert!(values.get(ApiKey::Language).as_text().is_some());
        assert!(values
            .get(ApiKey::Footer(FooterField::Bank))
            .as_number()
            .is_some());
        assert!(values
            .get(ApiKey::Position {
                category: 1,
                position: 1,
                field: PositionField::Approved
            })
            .as_number()
            .is_some());
    }
}
