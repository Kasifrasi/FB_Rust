//! Typsichere Typen für die Finanzbericht-API
//!
//! Dieses Modul stellt sicher, dass nur gültige Werte an die API übergeben werden können.
//! Ungültige Werte werden zur Compile-Zeit (Enums) oder zur Laufzeit mit klaren Fehlern abgefangen.

use crate::lang::data::{CURRENCIES, TEXT_MATRIX};
use std::fmt;
use std::str::FromStr;

// ============================================================================
// Language - Typsichere Sprachauswahl
// ============================================================================

/// Unterstützte Sprachen für den Finanzbericht
///
/// Diese Enum wird aus der TEXT_MATRIX generiert und stellt sicher,
/// dass nur gültige Sprachen verwendet werden können.
///
/// # Beispiel
/// ```ignore
/// use kmw_fb_rust::report::types::Language;
///
/// let lang = Language::Deutsch;
/// assert_eq!(lang.as_str(), "deutsch");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Language {
    /// Deutsch
    #[default]
    Deutsch,
    /// English
    English,
    /// Français
    Francais,
    /// Español
    Espanol,
    /// Português
    Portugues,
}

impl Language {
    /// Gibt den String-Wert zurück, wie er in TEXT_MATRIX verwendet wird
    pub fn as_str(&self) -> &'static str {
        match self {
            Language::Deutsch => "deutsch",
            Language::English => "english",
            Language::Francais => "français",
            Language::Espanol => "español",
            Language::Portugues => "português",
        }
    }

    /// Alle verfügbaren Sprachen
    pub fn all() -> &'static [Language] {
        &[
            Language::Deutsch,
            Language::English,
            Language::Francais,
            Language::Espanol,
            Language::Portugues,
        ]
    }

    /// Validiert, dass die Sprache in TEXT_MATRIX existiert
    pub fn validate(&self) -> bool {
        let lang_str = self.as_str();
        TEXT_MATRIX
            .iter()
            .any(|row| !row.is_empty() && row[0].eq_ignore_ascii_case(lang_str))
    }
}

impl fmt::Display for Language {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

impl FromStr for Language {
    type Err = String;

    /// Parses a language string (case-insensitive)
    ///
    /// Accepts various formats like "German", "de", "Deutsch", etc.
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let lower = s.to_lowercase();
        match lower.as_str() {
            "deutsch" | "german" | "de" => Ok(Language::Deutsch),
            "english" | "englisch" | "en" => Ok(Language::English),
            "français" | "francais" | "french" | "fr" => Ok(Language::Francais),
            "español" | "espanol" | "spanish" | "es" => Ok(Language::Espanol),
            "português" | "portugues" | "portuguese" | "pt" => Ok(Language::Portugues),
            _ => Err(format!("Unknown language: {}", s)),
        }
    }
}

// ============================================================================
// Currency - Validierte Währung
// ============================================================================

/// Validierte Währung (ISO 4217 Code)
///
/// Wird zur Laufzeit gegen CURRENCIES validiert, sodass Änderungen
/// an der Währungsliste automatisch berücksichtigt werden.
///
/// # Beispiel
/// ```ignore
/// use kmw_fb_rust::report::types::Currency;
///
/// let eur = Currency::new("EUR").unwrap();
/// let invalid = Currency::new("INVALID"); // None
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Currency(String);

impl Currency {
    /// Erstellt eine neue validierte Währung
    ///
    /// Gibt `None` zurück wenn der Code nicht in CURRENCIES existiert.
    /// Case-insensitive Suche, speichert aber immer in Uppercase.
    pub fn new(code: &str) -> Option<Self> {
        let upper = code.to_uppercase();
        if CURRENCIES.iter().any(|c| c.eq_ignore_ascii_case(&upper)) {
            Some(Currency(upper))
        } else {
            None
        }
    }

    /// Erstellt eine Währung ohne Validierung (für interne Nutzung)
    ///
    /// # Safety
    /// Nur verwenden wenn sicher ist, dass der Code gültig ist.
    pub fn new_unchecked(code: &str) -> Self {
        Currency(code.to_uppercase())
    }

    /// Gibt den Währungscode zurück
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Alle verfügbaren Währungen
    pub fn all() -> Vec<&'static str> {
        CURRENCIES.clone()
    }

    /// Prüft ob ein Währungscode gültig ist
    pub fn is_valid(code: &str) -> bool {
        CURRENCIES.iter().any(|c| c.eq_ignore_ascii_case(code))
    }

    // Häufig verwendete Währungen als Konstanten-artige Methoden
    pub fn eur() -> Self {
        Currency("EUR".to_string())
    }
    pub fn usd() -> Self {
        Currency("USD".to_string())
    }
    pub fn gbp() -> Self {
        Currency("GBP".to_string())
    }
    pub fn chf() -> Self {
        Currency("CHF".to_string())
    }
}

impl fmt::Display for Currency {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl Default for Currency {
    fn default() -> Self {
        Currency::eur()
    }
}

// ============================================================================
// Category - Kostenkategorien
// ============================================================================

/// Kostenkategorien des Finanzberichts
///
/// Diese Enum repräsentiert die 8 Hauptkategorien des Finanzberichts.
///
/// # Beispiel
/// ```ignore
/// use kmw_fb_rust::report::types::Category;
///
/// let config = BodyConfig::new()
///     .with_positions(Category::Personal, 10)
///     .with_positions(Category::Sachkosten, 15);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Category {
    /// Kategorie 1: Personalkosten
    Bauausgaben = 1,
    /// Kategorie 2: Sachkosten
    Investitionen = 2,
    /// Kategorie 3: Reisekosten
    Personalausgaben = 3,
    /// Kategorie 4: Investitionen
    Projektaktivitaeten = 4,
    /// Kategorie 5: Sonstige Kosten
    Projektverwaltung = 5,
    /// Kategorie 6: Projektverwaltung
    Evaluierung = 6,
    /// Kategorie 7: Evaluierung/Audit
    Audit = 7,
    /// Kategorie 8: Reserve
    Reserve = 8,
}

impl Category {
    /// Gibt den numerischen Index zurück (1-8)
    pub fn index(&self) -> u8 {
        *self as u8
    }

    /// Erstellt eine Kategorie aus einem Index
    ///
    /// Gibt `None` zurück wenn der Index nicht 1-8 ist.
    pub fn from_index(idx: u8) -> Option<Self> {
        match idx {
            1 => Some(Category::Bauausgaben),
            2 => Some(Category::Investitionen),
            3 => Some(Category::Personalausgaben),
            4 => Some(Category::Projektaktivitaeten),
            5 => Some(Category::Projektverwaltung),
            6 => Some(Category::Evaluierung),
            7 => Some(Category::Audit),
            8 => Some(Category::Reserve),
            _ => None,
        }
    }

    /// Alle Kategorien
    pub fn all() -> &'static [Category] {
        &[
            Category::Bauausgaben,
            Category::Investitionen,
            Category::Personalausgaben,
            Category::Projektaktivitaeten,
            Category::Projektverwaltung,
            Category::Evaluierung,
            Category::Audit,
            Category::Reserve,
        ]
    }

    /// Deutscher Name der Kategorie
    pub fn name_de(&self) -> &'static str {
        match self {
            Category::Bauausgaben => "Personalkosten",
            Category::Investitionen => "Sachkosten",
            Category::Personalausgaben => "Reisekosten",
            Category::Projektaktivitaeten => "Investitionen",
            Category::Projektverwaltung => "Sonstige Kosten",
            Category::Evaluierung => "Projektverwaltung",
            Category::Audit => "Evaluierung/Audit",
            Category::Reserve => "Reserve",
        }
    }

    /// Englischer Name der Kategorie
    pub fn name_en(&self) -> &'static str {
        match self {
            Category::Bauausgaben => "Personnel Costs",
            Category::Investitionen => "Material Costs",
            Category::Personalausgaben => "Travel Costs",
            Category::Projektaktivitaeten => "Investments",
            Category::Projektverwaltung => "Other Costs",
            Category::Evaluierung => "Project Administration",
            Category::Audit => "Evaluation/Audit",
            Category::Reserve => "Reserve",
        }
    }
}

impl fmt::Display for Category {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.name_de())
    }
}

// ============================================================================
// ReportDate - Validiertes Datum
// ============================================================================

/// Validiertes Datum für Finanzberichte
///
/// Stellt sicher, dass nur gültige Datumsangaben verwendet werden.
/// Unterstützt verschiedene Eingabeformate.
///
/// # Beispiel
/// ```ignore
/// use kmw_fb_rust::report::types::ReportDate;
///
/// let date = ReportDate::new(2024, 1, 15).unwrap();
/// assert_eq!(date.format_de(), "15.01.2024");
/// assert_eq!(date.format_iso(), "2024-01-15");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReportDate {
    year: u16,
    month: u8,
    day: u8,
}

/// Fehler bei der Datumserstellung
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DateError {
    /// Ungültiger Monat (muss 1-12 sein)
    InvalidMonth(u8),
    /// Ungültiger Tag für den gegebenen Monat
    InvalidDay { day: u8, month: u8, year: u16 },
    /// Ungültiges Jahr
    InvalidYear(u16),
    /// Ungültiges Format beim Parsen
    ParseError(String),
}

impl fmt::Display for DateError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DateError::InvalidMonth(m) => write!(f, "Ungültiger Monat: {} (muss 1-12 sein)", m),
            DateError::InvalidDay { day, month, year } => {
                write!(f, "Ungültiger Tag: {}.{}.{}", day, month, year)
            }
            DateError::InvalidYear(y) => write!(f, "Ungültiges Jahr: {}", y),
            DateError::ParseError(s) => write!(f, "Konnte Datum nicht parsen: {}", s),
        }
    }
}

impl std::error::Error for DateError {}

impl ReportDate {
    /// Erstellt ein neues validiertes Datum
    pub fn new(year: u16, month: u8, day: u8) -> Result<Self, DateError> {
        // Jahr-Validierung (sinnvoller Bereich für Finanzberichte)
        if !(1900..=2100).contains(&year) {
            return Err(DateError::InvalidYear(year));
        }

        // Monat-Validierung
        if !(1..=12).contains(&month) {
            return Err(DateError::InvalidMonth(month));
        }

        // Tag-Validierung
        let days_in_month = Self::days_in_month(year, month);
        if day < 1 || day > days_in_month {
            return Err(DateError::InvalidDay { day, month, year });
        }

        Ok(ReportDate { year, month, day })
    }

    /// Parst ein Datum aus verschiedenen Formaten
    ///
    /// Unterstützte Formate:
    /// - `DD.MM.YYYY` (deutsch)
    /// - `YYYY-MM-DD` (ISO)
    /// - `MM/DD/YYYY` (US)
    /// - `DD/MM/YYYY` (EU)
    pub fn parse(s: &str) -> Result<Self, DateError> {
        let s = s.trim();

        // Versuche verschiedene Formate
        if let Some(date) = Self::try_parse_de(s) {
            return Ok(date);
        }
        if let Some(date) = Self::try_parse_iso(s) {
            return Ok(date);
        }
        if let Some(date) = Self::try_parse_slash(s) {
            return Ok(date);
        }

        Err(DateError::ParseError(s.to_string()))
    }

    fn try_parse_de(s: &str) -> Option<Self> {
        // DD.MM.YYYY
        let parts: Vec<&str> = s.split('.').collect();
        if parts.len() != 3 {
            return None;
        }
        let day: u8 = parts[0].parse().ok()?;
        let month: u8 = parts[1].parse().ok()?;
        let year: u16 = parts[2].parse().ok()?;
        Self::new(year, month, day).ok()
    }

    fn try_parse_iso(s: &str) -> Option<Self> {
        // YYYY-MM-DD
        let parts: Vec<&str> = s.split('-').collect();
        if parts.len() != 3 {
            return None;
        }
        let year: u16 = parts[0].parse().ok()?;
        let month: u8 = parts[1].parse().ok()?;
        let day: u8 = parts[2].parse().ok()?;
        Self::new(year, month, day).ok()
    }

    fn try_parse_slash(s: &str) -> Option<Self> {
        // DD/MM/YYYY oder MM/DD/YYYY
        let parts: Vec<&str> = s.split('/').collect();
        if parts.len() != 3 {
            return None;
        }
        let first: u8 = parts[0].parse().ok()?;
        let second: u8 = parts[1].parse().ok()?;
        let year: u16 = parts[2].parse().ok()?;

        // Versuche DD/MM/YYYY zuerst
        if let Ok(date) = Self::new(year, second, first) {
            return Some(date);
        }
        // Dann MM/DD/YYYY
        Self::new(year, first, second).ok()
    }

    fn days_in_month(year: u16, month: u8) -> u8 {
        match month {
            1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
            4 | 6 | 9 | 11 => 30,
            2 => {
                if Self::is_leap_year(year) {
                    29
                } else {
                    28
                }
            }
            _ => 0,
        }
    }

    fn is_leap_year(year: u16) -> bool {
        (year.is_multiple_of(4) && !year.is_multiple_of(100)) || year.is_multiple_of(400)
    }

    /// Formatiert das Datum im deutschen Format (DD.MM.YYYY)
    pub fn format_de(&self) -> String {
        format!("{:02}.{:02}.{}", self.day, self.month, self.year)
    }

    /// Formatiert das Datum im ISO-Format (YYYY-MM-DD)
    pub fn format_iso(&self) -> String {
        format!("{}-{:02}-{:02}", self.year, self.month, self.day)
    }

    /// Formatiert das Datum im US-Format (MM/DD/YYYY)
    pub fn format_us(&self) -> String {
        format!("{:02}/{:02}/{}", self.month, self.day, self.year)
    }

    /// Getter für Jahr
    pub fn year(&self) -> u16 {
        self.year
    }

    /// Getter für Monat
    pub fn month(&self) -> u8 {
        self.month
    }

    /// Getter für Tag
    pub fn day(&self) -> u8 {
        self.day
    }
}

impl fmt::Display for ReportDate {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.format_de())
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_language_as_str() {
        assert_eq!(Language::Deutsch.as_str(), "deutsch");
        assert_eq!(Language::English.as_str(), "english");
        assert_eq!(Language::Francais.as_str(), "français");
    }

    #[test]
    fn test_language_from_str() {
        assert_eq!(Language::from_str("deutsch"), Ok(Language::Deutsch));
        assert_eq!(Language::from_str("German"), Ok(Language::Deutsch));
        assert_eq!(Language::from_str("english"), Ok(Language::English));
        assert_eq!(Language::from_str("englisch"), Ok(Language::English)); // Deutsche Schreibweise
        assert!(Language::from_str("invalid").is_err());
    }

    #[test]
    fn test_language_validate() {
        for lang in Language::all() {
            assert!(lang.validate(), "Language {:?} should be valid", lang);
        }
    }

    #[test]
    fn test_currency_new() {
        assert!(Currency::new("EUR").is_some());
        assert!(Currency::new("eur").is_some()); // Case-insensitive
        assert!(Currency::new("USD").is_some());
        assert!(Currency::new("INVALID").is_none());
        assert!(Currency::new("").is_none());
    }

    #[test]
    fn test_currency_is_valid() {
        assert!(Currency::is_valid("EUR"));
        assert!(Currency::is_valid("usd"));
        assert!(!Currency::is_valid("INVALID"));
    }

    #[test]
    fn test_category_index() {
        assert_eq!(Category::Bauausgaben.index(), 1);
        assert_eq!(Category::Reserve.index(), 8);
    }

    #[test]
    fn test_category_from_index() {
        assert_eq!(Category::from_index(1), Some(Category::Bauausgaben));
        assert_eq!(Category::from_index(8), Some(Category::Reserve));
        assert_eq!(Category::from_index(0), None);
        assert_eq!(Category::from_index(9), None);
    }

    #[test]
    fn test_date_new_valid() {
        assert!(ReportDate::new(2024, 1, 15).is_ok());
        assert!(ReportDate::new(2024, 2, 29).is_ok()); // Schaltjahr
        assert!(ReportDate::new(2023, 2, 28).is_ok());
    }

    #[test]
    fn test_date_new_invalid() {
        assert!(ReportDate::new(2024, 13, 1).is_err()); // Ungültiger Monat
        assert!(ReportDate::new(2024, 2, 30).is_err()); // Ungültiger Tag
        assert!(ReportDate::new(2023, 2, 29).is_err()); // Kein Schaltjahr
        assert!(ReportDate::new(1800, 1, 1).is_err()); // Jahr zu klein
    }

    #[test]
    fn test_date_parse() {
        // Deutsches Format
        let date = ReportDate::parse("15.01.2024").unwrap();
        assert_eq!(date.day(), 15);
        assert_eq!(date.month(), 1);
        assert_eq!(date.year(), 2024);

        // ISO Format
        let date = ReportDate::parse("2024-01-15").unwrap();
        assert_eq!(date.format_iso(), "2024-01-15");

        // Ungültig
        assert!(ReportDate::parse("invalid").is_err());
    }

    #[test]
    fn test_date_format() {
        let date = ReportDate::new(2024, 1, 15).unwrap();
        assert_eq!(date.format_de(), "15.01.2024");
        assert_eq!(date.format_iso(), "2024-01-15");
        assert_eq!(date.format_us(), "01/15/2024");
    }
}
