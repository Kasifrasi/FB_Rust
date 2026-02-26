use kmw_fb_rust::workbook_protection::protect_workbook;
use rayon::prelude::*;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Instant;

#[test]
fn test_parallel_protection_benchmark() -> anyhow::Result<()> {
    let file_count = 100;
    // Basis-Pfad (relativ zum Crate-Root)
    let output_dir = Path::new("tests/output");

    // 1. Eine existierende .xlsx Datei als Basis finden
    let base_file = fs::read_dir(output_dir)?
        .filter_map(|e| e.ok())
        .find(|e| e.path().extension().is_some_and(|ext| ext == "xlsx"))
        .map(|e| e.path());

    let base_file = match base_file {
        Some(f) => f,
        None => {
            println!(
                "Keine .xlsx Datei in {:?} gefunden. Test übersprungen.",
                output_dir
            );
            return Ok(());
        }
    };
    println!("Verwende Basis-Datei: {:?}", base_file);

    let benchmark_dir = output_dir.join("benchmark_input");
    let result_dir = output_dir.join("gesperrt");

    // Verzeichnisse vorbereiten
    if benchmark_dir.exists() {
        fs::remove_dir_all(&benchmark_dir)?;
    }
    fs::create_dir_all(&benchmark_dir)?;

    if result_dir.exists() {
        fs::remove_dir_all(&result_dir)?;
    }
    fs::create_dir_all(&result_dir)?;

    println!(
        "Erstelle {} Kopien von {:?} in {:?}...",
        file_count, base_file, benchmark_dir
    );
    let files: Vec<(PathBuf, PathBuf)> = (0..file_count)
        .map(|i| {
            let inp = benchmark_dir.join(format!("bench_{:03}.xlsx", i));
            let out = result_dir.join(format!("LOCKED_bench_{:03}.xlsx", i));
            fs::copy(&base_file, &inp).unwrap();
            (inp, out)
        })
        .collect();

    println!("Starte parallelen Schutz für {} Dateien...", file_count);
    let start = Instant::now();

    // 2. Parallele Verarbeitung mit Rayon
    files.par_iter().for_each(|(inp, out)| {
        if let Err(e) = protect_workbook(inp.to_str().unwrap(), out.to_str().unwrap(), "geheim123")
        {
            eprintln!("Fehler bei Datei {:?}: {}", inp, e);
        }
    });

    let duration = start.elapsed();
    let avg_per_file = duration.as_secs_f64() * 1000.0 / file_count as f64;
    let throughput = file_count as f64 / duration.as_secs_f64();

    println!("---------------------------------------------------");
    println!("BENCHMARK ERGEBNIS ({} Dateien, Parallel)", file_count);
    println!("Gesamtzeit:      {:.2?}", duration);
    println!("Durchschnitt:    {:.2} ms / Datei", avg_per_file);
    println!("Durchsatz:       {:.2} Dateien / Sekunde", throughput);
    println!("---------------------------------------------------");

    // Optional: Aufräumen des Inputs
    fs::remove_dir_all(benchmark_dir)?;

    Ok(())
}
