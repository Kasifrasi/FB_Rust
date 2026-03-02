//! Führt alle Report-Beispiele aus und kopiert die Ergebnisse nach /home/ardit/VM/Share.
//!
//! Usage: cargo run --example run_all_and_copy --release

use std::fs;
use std::path::Path;
use std::process::Command;

const OUTPUT_DIR: &str = "examples/output";
const TARGET_DIR: &str = "/home/ardit/VM/Share";

/// Alle Report-generierenden Beispiele (ohne profile/verify_password)
const EXAMPLES: &[&str] = &[
    "test_all_fields",
    "test_multilang",
    "test_multilang_grouped",
    "test_mixed_categories",
    "test_page_breaks",
    "test_workbook_protection",
    "test_parallel_protection",
];

fn main() {
    println!("=== Alle Beispiele ausführen und nach {} kopieren ===\n", TARGET_DIR);

    // 1. Bestehende .xlsx im Zielverzeichnis löschen
    let target = Path::new(TARGET_DIR);
    if target.exists() {
        let mut deleted = 0;
        for entry in fs::read_dir(target).expect("Kann Zielverzeichnis nicht lesen") {
            let entry = entry.expect("Fehler beim Lesen des Verzeichniseintrags");
            let path = entry.path();
            if path.extension().is_some_and(|ext| ext == "xlsx") {
                fs::remove_file(&path).expect("Kann Datei nicht löschen");
                deleted += 1;
            }
        }
        if deleted > 0 {
            println!("  {} bestehende .xlsx-Dateien in {} gelöscht.\n", deleted, TARGET_DIR);
        }
    } else {
        fs::create_dir_all(target).expect("Kann Zielverzeichnis nicht erstellen");
        println!("  Zielverzeichnis {} erstellt.\n", TARGET_DIR);
    }

    // 2. Alle Beispiele ausführen
    let mut failed = Vec::new();
    for example in EXAMPLES {
        print!("  Ausführen: {:<30} ", example);
        let output = Command::new("cargo")
            .args(["run", "--example", example, "--release"])
            .output()
            .expect("Kann cargo nicht ausführen");

        if output.status.success() {
            println!("OK");
        } else {
            println!("FEHLER");
            let stderr = String::from_utf8_lossy(&output.stderr);
            eprintln!("    {}", stderr.lines().last().unwrap_or("Unbekannter Fehler"));
            failed.push(*example);
        }
    }

    if !failed.is_empty() {
        eprintln!("\n  FEHLER bei: {}", failed.join(", "));
        std::process::exit(1);
    }

    // 3. Alle .xlsx aus examples/output/ nach Ziel kopieren
    let source = Path::new(OUTPUT_DIR);
    let mut copied = 0;
    for entry in fs::read_dir(source).expect("Kann Output-Verzeichnis nicht lesen") {
        let entry = entry.expect("Fehler beim Lesen");
        let path = entry.path();
        if path.extension().is_some_and(|ext| ext == "xlsx") {
            let dest = target.join(entry.file_name());
            fs::copy(&path, &dest).expect("Kann Datei nicht kopieren");
            copied += 1;
        }
    }

    println!("\n=== Fertig: {} Dateien nach {} kopiert ===", copied, TARGET_DIR);
}
