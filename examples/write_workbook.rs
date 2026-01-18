use kmw_fb_rust::{ExportConfig, FinanzberichtExporter};
use std::{fs, path::PathBuf};

fn main() -> anyhow::Result<()> {
    // Write demo workbooks for two languages into the repo-level output folder.
    let out_dir = PathBuf::from("../output");
    fs::create_dir_all(&out_dir)?;

    let exporter = FinanzberichtExporter::new();

    let mut cfg = ExportConfig::default();
    cfg.languages = vec!["Deutsch".to_string(), "Englisch".to_string()];
    cfg.output_folder = out_dir.clone();
    cfg.suffix = "_demo".to_string();
    cfg.protect = false; // skip protection for quick inspection

    let summary = exporter.export_all(&cfg)?;
    for path in summary.files {
        println!("wrote {}", path.display());
    }

    Ok(())
}