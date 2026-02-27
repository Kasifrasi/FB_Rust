use once_cell::sync::Lazy;
use std::collections::HashMap;

#[derive(Debug, Clone)]
pub struct LanguageConfig {
    pub base: &'static str,
    pub lang_val: &'static str,
    pub fb_sheet: &'static str,
    pub decimal_err: &'static str,
}

pub static LANG_CONFIG: Lazy<HashMap<&'static str, LanguageConfig>> = Lazy::new(|| {
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

pub static LANG_SUFFIXES: Lazy<HashMap<&'static str, &'static str>> = Lazy::new(|| {
    HashMap::from([
        ("Deutsch", "_de"),
        ("Englisch", "_en"),
        ("Französisch", "_fr"),
        ("Spanisch", "_es"),
        ("Portugiesisch", "_po"),
    ])
});
