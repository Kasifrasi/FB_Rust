//! Typsichere Typen für die Finanzbericht-API
//!
//! Dieses Modul stellt sicher, dass nur gültige Werte an die API übergeben werden können.
//! Ungültige Werte werden zur Compile-Zeit (Enums) oder zur Laufzeit mit klaren Fehlern abgefangen.

use crate::lang::data::TEXT_MATRIX;
use std::fmt;
use std::str::FromStr;

// ============================================================================
// Language - Typsichere Sprachauswahl
// ============================================================================

/// Supported languages for the financial report.
///
/// This enum ensures only valid languages can be used. The available
/// languages are derived from the internal text matrix.
///
/// # Example
/// ```ignore
/// use fb_rust::Language;
///
/// let lang = Language::Deutsch;
/// assert_eq!(lang.as_str(), "deutsch");
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Serialized with `rename_all = "lowercase"`:
///
/// ```json
/// "deutsch"
/// "english"
/// "francais"
/// "espanol"
/// "portugues"
/// ```
///
/// [`FromStr`] accepts additional aliases (`"german"`, `"de"`, `"fr"`, etc.),
/// but serde only accepts the exact lowercase variant names above.
///
/// Default: `Deutsch`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "lowercase"))]
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
// Currency - ISO 4217 Enum (defined in lang::data, re-exported here)
// ============================================================================

pub use crate::lang::data::Currency;

// ============================================================================
// Category - Kostenkategorien
// ============================================================================

/// Cost categories of the financial report.
///
/// Represents the 8 main categories. Used with [`BodyConfig::with_cat_positions`](crate::BodyConfig::with_cat_positions)
/// for the typed Rust API.
///
/// # Example
/// ```ignore
/// use fb_rust::{BodyConfig, Category};
///
/// let config = BodyConfig::new()
///     .with_cat_positions(Category::Bauausgaben, 10)
///     .with_cat_positions(Category::Evaluierung, 0);
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Serialized using PascalCase variant names:
///
/// | JSON value | Index | Description |
/// |---|---|---|
/// | `"Bauausgaben"` | 1 | Personnel costs |
/// | `"Investitionen"` | 2 | Material costs |
/// | `"Personalausgaben"` | 3 | Travel costs |
/// | `"Projektaktivitaeten"` | 4 | Investments |
/// | `"Projektverwaltung"` | 5 | Other costs |
/// | `"Evaluierung"` | 6 | Project administration |
/// | `"Audit"` | 7 | Evaluation/Audit |
/// | `"Reserve"` | 8 | Reserve |
///
/// **Note:** In [`ReportConfig`](crate::ReportConfig) JSON, categories are referenced
/// by numeric keys (`"1"` – `"8"`) as keys of the `positions` and `header_inputs` maps.
/// The `Category` enum is for the typed Rust API only.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
pub enum Category {
    /// Kategorie 1: Bauausgaben
    Bauausgaben = 1,
    /// Kategorie 2: Investitionen
    Investitionen = 2,
    /// Kategorie 3: Personalausgaben
    Personalausgaben = 3,
    /// Kategorie 4: Projektaktivitäten
    Projektaktivitaeten = 4,
    /// Kategorie 5: Projektverwaltung
    Projektverwaltung = 5,
    /// Kategorie 6: Evaluierung
    Evaluierung = 6,
    /// Kategorie 7: Audit
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

/// Validated date for financial reports.
///
/// Ensures only valid dates are used. Supports multiple input formats.
/// Valid year range: 1900–2100.
///
/// # Example
/// ```ignore
/// use fb_rust::ReportDate;
///
/// let date = ReportDate::new(2024, 1, 15).unwrap();
/// assert_eq!(date.format_de(), "15.01.2024");
/// assert_eq!(date.format_iso(), "2024-01-15");
/// ```
///
/// ## Serde (requires `serde` feature)
///
/// Deserialized via `try_from = "String"` using [`ReportDate::parse`], which accepts:
/// - `"2024-06-15"` — ISO 8601
/// - `"15.06.2024"` — German (DD.MM.YYYY)
/// - `"15/06/2024"` — EU (DD/MM/YYYY)
/// - `"06/15/2024"` — US (MM/DD/YYYY, tried as fallback)
///
/// Always serialized to ISO format:
/// ```json
/// "2024-06-15"
/// ```
///
/// **Note:** In [`ReportConfig`](crate::ReportConfig), date fields (`project_start`,
/// `report_end`, etc.) are plain `Option<String>`, not `ReportDate`. The `ReportDate`
/// type is for the typed Rust API via [`ReportValues`](crate::ReportValues).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg_attr(feature = "serde", derive(serde::Serialize, serde::Deserialize))]
#[cfg_attr(feature = "serde", serde(try_from = "String", into = "String"))]
pub struct ReportDate {
    year: u16,
    month: u8,
    day: u8,
}

/// Error from date construction.
///
/// Returned by [`ReportDate::new`] and [`ReportDate::parse`].
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

#[cfg(feature = "serde")]
impl From<ReportDate> for String {
    fn from(d: ReportDate) -> Self {
        d.format_iso()
    }
}

#[cfg(feature = "serde")]
impl TryFrom<String> for ReportDate {
    type Error = DateError;
    fn try_from(s: String) -> Result<Self, Self::Error> {
        Self::parse(&s)
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
    fn test_currency_parse() {
        assert_eq!("EUR".parse::<Currency>(), Ok(Currency::EUR));
        assert_eq!("eur".parse::<Currency>(), Ok(Currency::EUR)); // case-insensitive
        assert_eq!("USD".parse::<Currency>(), Ok(Currency::USD));
        assert!("INVALID".parse::<Currency>().is_err());
        assert!("".parse::<Currency>().is_err());
    }

    #[test]
    fn test_currency_as_str() {
        assert_eq!(Currency::EUR.as_str(), "EUR");
        assert_eq!(Currency::USD.as_str(), "USD");
        assert_eq!(Currency::GBP.as_str(), "GBP");
    }

    #[test]
    fn test_currency_all() {
        let all = Currency::all();
        assert_eq!(all.len(), 153);
        assert!(all.contains(&Currency::EUR));
        assert!(all.contains(&Currency::USD));
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
