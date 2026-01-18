use crate::dynamic_report::build_dynamic_report;
use crate::formats::StylePalette;
use crate::language::{
    CURRENCIES, DEFAULT_NUM_POSITIONS, LANG_CONFIG, LANG_SUFFIXES, PROTECTION_DEFAULTS, TEXT_MATRIX,
};
use anyhow::{bail, Context, Result};
use rust_xlsxwriter::{ProtectionOptions, Workbook, Worksheet};
use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Debug, Clone)]
pub struct ExportConfig {
    pub languages: Vec<String>,
    pub output_folder: PathBuf,
    pub suffix: String,
    pub num_positions: HashMap<u8, u16>,
    pub protect: bool,
    pub password: String,
    pub protection_overrides: HashMap<String, bool>,
}

impl Default for ExportConfig {
    fn default() -> Self {
        Self {
            languages: LANG_CONFIG.keys().map(|lang| (*lang).to_string()).collect(),
            output_folder: PathBuf::from("output"),
            suffix: String::new(),
            num_positions: DEFAULT_NUM_POSITIONS.clone(),
            protect: true,
            password: "thetool".to_string(),
            protection_overrides: HashMap::new(),
        }
    }
}

#[derive(Debug, Default, Clone)]
pub struct ExportSummary {
    pub files: Vec<PathBuf>,
}

pub struct FinanzberichtExporter {
    palette: StylePalette,
}

impl FinanzberichtExporter {
    pub fn new() -> Self {
        Self {
            palette: StylePalette::new(),
        }
    }

    pub fn export_all(&self, config: &ExportConfig) -> Result<ExportSummary> {
        if !config.output_folder.exists() {
            bail!(
                "Output folder does not exist: {}",
                config.output_folder.display()
            );
        }

        let mut summary = ExportSummary::default();

        // Ensure translation sheet exists before per-language sheets so VLOOKUP targets are present.
        // We re-create it for each workbook to keep the API simple.

        for language in &config.languages {
            let mut workbook = self.create_workbook(language, config)?;
            let cfg = LANG_CONFIG
                .get(language.as_str())
                .with_context(|| format!("Unsupported language: {language}"))?;

            let filename = format!("{}{}{}.xlsx", cfg.base, config.suffix, "");
            let path = config.output_folder.join(filename);
            workbook
                .save(path.to_str().context("Invalid UTF-8 in output path")?)
                .with_context(|| "Failed to write Excel workbook")?;

            summary.files.push(path);
        }

        Ok(summary)
    }

    fn create_workbook(&self, language: &str, config: &ExportConfig) -> Result<Workbook> {
        let cfg = LANG_CONFIG
            .get(language)
            .with_context(|| format!("Unsupported language: {language}"))?;

        let mut workbook = Workbook::new();

        // Hidden translation sheet used by VLOOKUPs
        self.build_translation_sheet(&mut workbook)?;

        crate::sheet_builder::recreate_sheet(&mut workbook, cfg.fb_sheet)
            .map_err(|e| anyhow::anyhow!(e))
            .context("Failed to recreate sheet template")?;

        let worksheet = workbook
            .worksheet_from_name(cfg.fb_sheet)
            .map_err(|e| anyhow::anyhow!(e))
            .context("Failed to get worksheet")?;

        // Ensure cached formula results default to empty so viewers without recalc don't show 0.
        worksheet.set_formula_result_default("");

        self.apply_header(worksheet, cfg, language, config)?;

        Ok(workbook)
    }

    fn apply_header(
        &self,
        worksheet: &mut Worksheet,
        cfg: &crate::language::LanguageConfig,
        language: &str,
        config: &ExportConfig,
    ) -> Result<()> {
        let suffix = LANG_SUFFIXES.get(language).copied().unwrap_or("");

        worksheet.write_string(1, 1, suffix)?;
        worksheet.write_string(1, 4, cfg.lang_val)?;

        let last_row = build_dynamic_report(worksheet, &self.palette, &config.num_positions)?;

        if config.protect {
            self.apply_protection(worksheet, config, last_row)?;
        }

        Ok(())
    }

    fn build_translation_sheet(&self, workbook: &mut Workbook) -> Result<()> {
        let ws = workbook.add_worksheet().set_name("Sprachversionen")?;

        for (row_idx, currency) in CURRENCIES.iter().enumerate() {
            ws.write_string(row_idx as u32, 0, *currency)?;
        }

        for (row_idx, row) in TEXT_MATRIX.iter().enumerate() {
            for (col_idx, value) in row.iter().enumerate() {
                ws.write_string(row_idx as u32, (col_idx + 1) as u16, *value)?;
            }
        }

        ws.set_hidden(true);
        Ok(())
    }

    fn apply_protection(
        &self,
        worksheet: &mut Worksheet,
        config: &ExportConfig,
        _last_row: u32,
    ) -> Result<()> {
        let defaults = PROTECTION_DEFAULTS.to_flags();
        let mut options = ProtectionOptions::new();

        for (key, default_value) in defaults {
            let value = config
                .protection_overrides
                .get(key)
                .copied()
                .unwrap_or(default_value);
            match key {
                "AllowFormattingCells" => options.format_cells = value,
                "AllowFormattingColumns" => options.format_columns = value,
                "AllowFormattingRows" => options.format_rows = value,
                "AllowInsertingColumns" => options.insert_columns = value,
                "AllowInsertingRows" => options.insert_rows = value,
                "AllowDeletingColumns" => options.delete_columns = value,
                "AllowDeletingRows" => options.delete_rows = value,
                "AllowSorting" => options.sort = value,
                "AllowFiltering" => options.use_autofilter = value,
                "AllowUsingPivotTables" => options.use_pivot_tables = value,
                "AllowInsertingHyperlinks" => options.insert_links = value,
                "AllowDrawingObjects" => options.edit_objects = value,
                "AllowScenarios" => options.edit_scenarios = value,
                _ => {}
            }
        }

        worksheet.protect_with_options(&options);

        if !config.password.is_empty() {
            worksheet.protect_with_password(&config.password);
        }

        Ok(())
    }
}
