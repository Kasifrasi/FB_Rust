use once_cell::sync::Lazy;
use std::collections::HashMap;

#[derive(Debug, Clone)]
pub struct LanguageConfig {
    pub base: &'static str,
    pub lang_val: &'static str,
    pub fb_sheet: &'static str,
    pub rename: &'static [(&'static str, &'static str)],
    pub decimal_err: &'static str,
}

#[derive(Debug, Clone, Copy)]
pub struct ProtectionDefaults {
    pub allow_formatting_cells: bool,
    pub allow_formatting_columns: bool,
    pub allow_formatting_rows: bool,
    pub allow_inserting_columns: bool,
    pub allow_inserting_rows: bool,
    pub allow_deleting_columns: bool,
    pub allow_deleting_rows: bool,
    pub allow_sorting: bool,
    pub allow_filtering: bool,
    pub allow_using_pivot_tables: bool,
    pub allow_inserting_hyperlinks: bool,
    pub allow_drawing_objects: bool,
    pub allow_scenarios: bool,
}

impl ProtectionDefaults {
    pub fn to_flags(self) -> HashMap<&'static str, bool> {
        HashMap::from([
            ("AllowFormattingCells", self.allow_formatting_cells),
            ("AllowFormattingColumns", self.allow_formatting_columns),
            ("AllowFormattingRows", self.allow_formatting_rows),
            ("AllowInsertingColumns", self.allow_inserting_columns),
            ("AllowInsertingRows", self.allow_inserting_rows),
            ("AllowDeletingColumns", self.allow_deleting_columns),
            ("AllowDeletingRows", self.allow_deleting_rows),
            ("AllowSorting", self.allow_sorting),
            ("AllowFiltering", self.allow_filtering),
            ("AllowUsingPivotTables", self.allow_using_pivot_tables),
            ("AllowInsertingHyperlinks", self.allow_inserting_hyperlinks),
            ("AllowDrawingObjects", self.allow_drawing_objects),
            ("AllowScenarios", self.allow_scenarios),
        ])
    }
}

pub static LANG_CONFIG: Lazy<HashMap<&'static str, LanguageConfig>> = Lazy::new(|| {
    HashMap::from([
        (
            "Deutsch",
            LanguageConfig {
                base: "D_Finanzbericht",
                lang_val: "deutsch",
                fb_sheet: "Finanzbericht",
                rename: &[],
                decimal_err: "Bitte geben Sie maximal 2 Nachkommastellen ein.",
            },
        ),
        (
            "Englisch",
            LanguageConfig {
                base: "E_Financial_report",
                lang_val: "english",
                fb_sheet: "Financial Report",
                rename: &[("Finanzbericht", "Financial Report")],
                decimal_err: "Please enter a maximum of 2 decimal places.",
            },
        ),
        (
            "Französisch",
            LanguageConfig {
                base: "F_Rapport_financier",
                lang_val: "français",
                fb_sheet: "Rapport financier",
                rename: &[("Finanzbericht", "Rapport financier")],
                decimal_err: "Veuillez saisir un maximum de 2 décimales.",
            },
        ),
        (
            "Spanisch",
            LanguageConfig {
                base: "S_Informe_financiero",
                lang_val: "español",
                fb_sheet: "Informe financiero",
                rename: &[("Finanzbericht", "Informe financiero")],
                decimal_err: "Por favor, introduzca un máximo de 2 decimales.",
            },
        ),
        (
            "Portugiesisch",
            LanguageConfig {
                base: "P_Relatorio_financeiro",
                lang_val: "português",
                fb_sheet: "Relatório financeiro",
                rename: &[("Finanzbericht", "Relatório financeiro")],
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

pub static DEFAULT_NUM_POSITIONS: Lazy<HashMap<u8, u16>> =
    Lazy::new(|| HashMap::from([(1, 20), (2, 20), (3, 30), (4, 30), (5, 20)]));

pub static PROTECTION_DEFAULTS: ProtectionDefaults = ProtectionDefaults {
    allow_formatting_cells: true,
    allow_formatting_columns: true,
    allow_formatting_rows: true,
    allow_inserting_columns: false,
    allow_inserting_rows: false,
    allow_deleting_columns: false,
    allow_deleting_rows: false,
    allow_sorting: true,
    allow_filtering: true,
    allow_using_pivot_tables: true,
    allow_inserting_hyperlinks: true,
    allow_drawing_objects: false,
    allow_scenarios: true,
};
