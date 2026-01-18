mod formats;
mod workbook;
mod language;
mod dynamic_report;
mod sheet_builder;

pub use crate::formats::StylePalette;
pub use crate::workbook::{ExportConfig, ExportSummary, FinanzberichtExporter};

#[cfg(test)]
mod tests {
	use super::*;
	use std::fs;

	#[test]
	fn export_writes_workbook_file() {
		let tmp_dir = tempfile::tempdir().expect("tempdir");
		let out_dir = tmp_dir.path().join("out");
		fs::create_dir(&out_dir).expect("create out dir");

		let exporter = FinanzberichtExporter::new();
		let mut config = ExportConfig::default();
		config.languages = vec!["Deutsch".to_string()];
		config.output_folder = out_dir.clone();
		config.protect = false; // simplifies smoke test

		let summary = exporter.export_all(&config).expect("export succeeds");
		assert_eq!(summary.files.len(), 1, "one workbook produced");

		let workbook_path = &summary.files[0];
		let meta = fs::metadata(workbook_path).expect("workbook written");
		assert!(meta.is_file(), "output is a file");
		assert!(meta.len() > 0, "output file not empty");
	}
}
