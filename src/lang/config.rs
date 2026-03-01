//! Sprachkonfiguration für alle unterstützten Sprachen
//!
//! Definiert [`LanguageConfig`] mit den sprachspezifischen Sheet-Namen und
//! Fehlermeldungen sowie die statischen Maps [`LANG_CONFIG`] und [`LANG_SUFFIXES`].

use std::collections::HashMap;
use std::sync::LazyLock;

/// Sprachspezifische Konfiguration für Sheet-Namen und Fehlermeldungen
#[derive(Debug, Clone)]
pub struct LanguageConfig {
    /// Basisname für generierte Dateinamen (z.B. `"D_Finanzbericht"`)
    pub base: &'static str,
    /// Sprachcode wie er in TEXT_MATRIX steht (z.B. `"deutsch"`)
    pub lang_val: &'static str,
    /// Lokalisierter Sheet-Name (z.B. `"Finanzbericht"`)
    pub fb_sheet: &'static str,
    /// Fehlermeldung für Dezimalstellen-Validierung in der jeweiligen Sprache
    pub decimal_err: &'static str,
}

/// Sprachkonfiguration, indiziert nach deutschem Sprachnamen (z.B. `"Deutsch"`, `"Englisch"`)
pub static LANG_CONFIG: LazyLock<HashMap<&'static str, LanguageConfig>> = LazyLock::new(|| {
    HashMap::from([
        (
            "Deutsch",
            LanguageConfig {
                base: "D_Finanzbericht",
                lang_val: "deutsch",
                fb_sheet: "Finanzbericht",
                decimal_err: "Bitte geben Sie maximal 2 Nachkommastellen ein.",
            },
        ),
        (
            "Englisch",
            LanguageConfig {
                base: "E_Financial_report",
                lang_val: "english",
                fb_sheet: "Financial Report",
                decimal_err: "Please enter a maximum of 2 decimal places.",
            },
        ),
        (
            "Französisch",
            LanguageConfig {
                base: "F_Rapport_financier",
                lang_val: "français",
                fb_sheet: "Rapport financier",
                decimal_err: "Veuillez saisir un maximum de 2 décimales.",
            },
        ),
        (
            "Spanisch",
            LanguageConfig {
                base: "S_Informe_financiero",
                lang_val: "español",
                fb_sheet: "Informe financiero",
                decimal_err: "Por favor, introduzca un máximo de 2 decimales.",
            },
        ),
        (
            "Portugiesisch",
            LanguageConfig {
                base: "P_Relatorio_financeiro",
                lang_val: "português",
                fb_sheet: "Relatório financeiro",
                decimal_err: "Por favor, insira no máximo 2 casas decimais.",
            },
        ),
    ])
});

/// Dateiname-Suffixe pro Sprache (z.B. `"Deutsch"` → `"_de"`)
pub static LANG_SUFFIXES: LazyLock<HashMap<&'static str, &'static str>> = LazyLock::new(|| {
    HashMap::from([
        ("Deutsch", "_de"),
        ("Englisch", "_en"),
        ("Französisch", "_fr"),
        ("Spanisch", "_es"),
        ("Portugiesisch", "_po"),
    ])
});
