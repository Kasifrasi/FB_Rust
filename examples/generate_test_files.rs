use kmw_fb_rust::{ExportConfig, FinanzberichtExporter};
use std::fs;
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    println!("========================================");
    println!("Starting Test File Generation");
    println!("========================================");

    // 1. Ensure output directory exists
    let output_dir = PathBuf::from("Test-Dateien");
    if !output_dir.exists() {
        println!("Creating directory: {:?}", output_dir);
        fs::create_dir(&output_dir)?;
    } else {
        println!("Using existing directory: {:?}", output_dir);
    }

    let exporter = FinanzberichtExporter::new();

    // ---------------------------------------------------------
    // Scenario 1: Standard German Report (Unprotected)
    // ---------------------------------------------------------
    println!("\n[1/4] Generating Standard German Report (Unprotected)...");
    let mut config_de = ExportConfig::default();
    config_de.output_folder = output_dir.clone();
    config_de.languages = vec!["Deutsch".to_string()];
    config_de.suffix = "_TEST_DE_Unprotected".to_string();
    config_de.protect = false;

    match exporter.export_all(&config_de) {
        Ok(summary) => {
            for file in summary.files {
                println!("  ✅ Created: {:?}", file);
            }
        }
        Err(e) => println!("  ❌ Error: {}", e),
    }

    // ---------------------------------------------------------
    // Scenario 2: English Report (Protected with password 'test')
    // ---------------------------------------------------------
    println!("\n[2/4] Generating English Report (Protected)...");
    let mut config_en = ExportConfig::default();
    config_en.output_folder = output_dir.clone();
    config_en.languages = vec!["Englisch".to_string()];
    config_en.suffix = "_TEST_EN_Protected".to_string();
    config_en.protect = true;
    config_en.password = "test".to_string();

    match exporter.export_all(&config_en) {
        Ok(summary) => {
            for file in summary.files {
                println!("  ✅ Created: {:?}", file);
            }
        }
        Err(e) => println!("  ❌ Error: {}", e),
    }

    // ---------------------------------------------------------
    // Scenario 3: Custom Row Counts (Dynamic Size)
    // ---------------------------------------------------------
    println!("\n[3/4] Generating Report with Custom Row Counts...");
    let mut config_custom = ExportConfig::default();
    config_custom.output_folder = output_dir.clone();
    config_custom.languages = vec!["Deutsch".to_string()];
    config_custom.suffix = "_TEST_CustomRows".to_string();
    config_custom.protect = false;

    // Modify default row counts for categories
    // Category 1: very short (5 rows)
    // Category 2: very long (30 rows)
    if let Some(count) = config_custom.num_positions.get_mut(&1) {
        *count = 5;
    }
    if let Some(count) = config_custom.num_positions.get_mut(&2) {
        *count = 30;
    }

    match exporter.export_all(&config_custom) {
        Ok(summary) => {
            for file in summary.files {
                println!("  ✅ Created: {:?}", file);
            }
        }
        Err(e) => println!("  ❌ Error: {}", e),
    }

    // ---------------------------------------------------------
    // Scenario 4: Multi-Language Batch
    // ---------------------------------------------------------
    println!("\n[4/4] Generating All Languages Batch...");
    let mut config_all = ExportConfig::default();
    config_all.output_folder = output_dir.clone();
    // Use all available languages defined in default config
    config_all.suffix = "_TEST_Batch".to_string();
    config_all.protect = false;

    match exporter.export_all(&config_all) {
        Ok(summary) => {
            for file in summary.files {
                println!("  ✅ Created: {:?}", file);
            }
        }
        Err(e) => println!("  ❌ Error: {}", e),
    }

    println!("\n========================================");
    println!("Generation Complete.");
    println!("Check the folder '{:?}' for results.", output_dir);
    println!("========================================");

    Ok(())
}
