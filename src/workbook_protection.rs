use anyhow::{Context, Result};
use base64::{engine::general_purpose, Engine as _};
use byteorder::{WriteBytesExt, LE};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use sha2::{Digest, Sha512};
use std::fs::File;
use std::io::{Cursor, Read, Write};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const DEFAULT_SPIN_COUNT: u32 = 100_000;
const SALT_SIZE: usize = 16;

/// Fester Salt für deterministische Hash-Berechnung.
/// Gleicher Salt → gleiches Passwort → gleicher Hash (kein Zufallselement).
const FIXED_SALT: [u8; 16] = [
    0x6b, 0x6d, 0x77, 0x66, 0x62, 0x5f, 0x72, 0x70, // "kmwfb_rp"
    0x74, 0x5f, 0x76, 0x31, 0x5f, 0x21, 0x21, 0x00, // "t_v1_!!\0"
];

/// Vorberechneter Workbook-Protection-Hash.
///
/// Einmalig mit [`precompute_hash`] oder [`precompute_hash_with_spin_count`] erstellen
/// und dann für beliebig viele Dateien mit [`protect_workbook_precomputed`] wiederverwenden.
/// Spart bei Batch-Operationen den teuren SHA-512-Iterationsaufwand (≈25 ms / 100.000 Iterationen).
pub struct PrecomputedHash {
    pub salt_b64: String,
    pub hash_b64: String,
    pub spin_count: u32,
}

/// Schneller spin count (ausreichend gegen Gelegenheitsnutzer)
const FAST_SPIN_COUNT: u32 = 1_000;

/// Workbook-level protection settings
///
/// Protects the workbook structure (prevents adding, deleting, moving, or renaming sheets).
/// Uses SHA-512 based password hashing according to ECMA-376 standard.
///
/// # Example
/// ```ignore
/// use kmw_fb_rust::WorkbookProtection;
///
/// // Standard (100.000 Iterationen, ~25ms pro Report)
/// let protection = WorkbookProtection::new("secret123");
///
/// // Schnell (1.000 Iterationen, ~0.3ms pro Report)
/// let fast = WorkbookProtection::fast("secret123");
///
/// // Benutzerdefiniert
/// let custom = WorkbookProtection::new("secret123").with_spin_count(10_000);
/// ```
#[derive(Debug, Clone)]
pub struct WorkbookProtection {
    /// Password for workbook protection
    pub password: String,
    /// Lock structure (prevent sheet add/delete/move/rename)
    pub lock_structure: bool,
    /// SHA-512 Iteration count (higher = slower but more secure)
    ///
    /// - `100_000` (Standard): ECMA-376 konform, ~25ms pro Report
    /// - `1_000` (Fast): Hält Gelegenheitsnutzer ab, ~0.3ms pro Report
    pub spin_count: u32,
}

impl WorkbookProtection {
    /// Creates new workbook protection with the given password
    ///
    /// Uses the standard ECMA-376 spin count of 100.000 iterations.
    pub fn new(password: impl Into<String>) -> Self {
        Self {
            password: password.into(),
            lock_structure: true,
            spin_count: DEFAULT_SPIN_COUNT,
        }
    }

    /// Creates fast workbook protection (1.000 iterations instead of 100.000)
    ///
    /// ~80x schneller als `new()`. Ausreichend um Gelegenheitsnutzer abzuhalten,
    /// aber nicht sicher gegen gezielte Brute-Force-Angriffe.
    pub fn fast(password: impl Into<String>) -> Self {
        Self {
            password: password.into(),
            lock_structure: true,
            spin_count: FAST_SPIN_COUNT,
        }
    }

    /// Sets a custom spin count for SHA-512 hashing
    pub fn with_spin_count(mut self, count: u32) -> Self {
        self.spin_count = count;
        self
    }

    /// Sets whether to lock workbook structure
    pub fn lock_structure(mut self, lock: bool) -> Self {
        self.lock_structure = lock;
        self
    }
}

/// Berechnet den Workbook-Protection-Hash einmalig mit festem Salt (100.000 Iterationen).
///
/// Der Hash ist deterministisch: gleiches Passwort → immer gleicher Hash.
/// Für Batch-Schutz vieler Dateien mit demselben Passwort einmalig aufrufen
/// und das Ergebnis an [`protect_workbook_precomputed`] übergeben.
pub fn precompute_hash(password: &str) -> PrecomputedHash {
    precompute_hash_with_spin_count(password, DEFAULT_SPIN_COUNT)
}

/// Berechnet den Workbook-Protection-Hash einmalig mit festem Salt und konfigurierbarem Spin Count.
pub fn precompute_hash_with_spin_count(password: &str, spin_count: u32) -> PrecomputedHash {
    let hash_b64 = compute_hash(password, &FIXED_SALT, spin_count);
    PrecomputedHash {
        salt_b64: general_purpose::STANDARD.encode(FIXED_SALT),
        hash_b64,
        spin_count,
    }
}

/// Schützt eine Workbook-Struktur mit einem vorberechneten Hash (kein Hash-Aufwand).
///
/// Deutlich schneller als [`protect_workbook`] bei Batch-Operationen, da der
/// SHA-512-Hash nicht neu berechnet wird, sondern direkt in die XML-Datei injiziert wird.
///
/// # Argumente
/// * `input_path` - Pfad zur ungeschützten .xlsx Datei
/// * `output_path` - Pfad, wo die geschützte Datei gespeichert werden soll
/// * `hash` - Vorberechneter Hash aus [`precompute_hash`]
pub fn protect_workbook_precomputed(
    input_path: &str,
    output_path: &str,
    hash: &PrecomputedHash,
) -> Result<()> {
    write_protected_zip(input_path, output_path, &hash.salt_b64, &hash.hash_b64, hash.spin_count)
}

/// Schützt die Arbeitsmappenstruktur einer existierenden Excel-Datei mit einem Passwort.
///
/// Dies verhindert, dass Benutzer Blätter hinzufügen, löschen, verschieben oder umbenennen.
/// Die Implementierung folgt dem ECMA-376 Standard (Agile Encryption mit SHA-512).
///
/// Verwendet den Standard spin count von 100.000 Iterationen.
///
/// # Argumente
/// * `input_path` - Pfad zur ungeschützten .xlsx Datei
/// * `output_path` - Pfad, wo die geschützte Datei gespeichert werden soll
/// * `password` - Das zu setzende Passwort
pub fn protect_workbook(input_path: &str, output_path: &str, password: &str) -> Result<()> {
    protect_workbook_with_spin_count(input_path, output_path, password, DEFAULT_SPIN_COUNT)
}

/// Schützt die Arbeitsmappenstruktur mit konfigurierbarem spin count.
///
/// # Argumente
/// * `input_path` - Pfad zur ungeschützten .xlsx Datei
/// * `output_path` - Pfad, wo die geschützte Datei gespeichert werden soll
/// * `password` - Das zu setzende Passwort
/// * `spin_count` - Anzahl SHA-512 Iterationen (höher = sicherer aber langsamer)
pub fn protect_workbook_with_spin_count(
    input_path: &str,
    output_path: &str,
    password: &str,
    spin_count: u32,
) -> Result<()> {
    let (salt_b64, hash_b64) = hash_password(password, spin_count);
    write_protected_zip(input_path, output_path, &salt_b64, &hash_b64, spin_count)
}

/// Kernfunktion: ZIP lesen, workbook.xml mit Protection-Tag versehen, neu schreiben.
///
/// Alle Einträge außer `xl/workbook.xml` werden per `raw_copy_file` 1:1 kopiert –
/// ohne Dekomprimierung oder Rekomprimierung. Nur workbook.xml wird gelesen,
/// modifiziert und neu komprimiert.
fn write_protected_zip(
    input_path: &str,
    output_path: &str,
    salt_b64: &str,
    hash_b64: &str,
    spin_count: u32,
) -> Result<()> {
    let file = File::open(input_path).context("Konnte Input-Datei nicht öffnen")?;
    let mut archive = ZipArchive::new(file).context("Konnte ZIP-Archiv nicht lesen")?;

    let out_file = File::create(output_path).context("Konnte Output-Datei nicht erstellen")?;
    let mut zip_writer = ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let name = archive.by_index(i)?.name().to_string();

        if name == "xl/workbook.xml" {
            let mut file = archive.by_index(i)?;
            let compression = file.compression();
            let unix_mode = file.unix_mode();
            let mut content = Vec::new();
            file.read_to_end(&mut content)?;

            let new_xml = inject_protection(&content, salt_b64, hash_b64, spin_count)?;
            let options = FileOptions::<()>::default()
                .compression_method(compression)
                .unix_permissions(unix_mode.unwrap_or(0o644));
            zip_writer.start_file(&name, options)?;
            zip_writer.write_all(&new_xml)?;
        } else {
            let file = archive.by_index(i)?;
            zip_writer.raw_copy_file(file)?;
        }
    }

    zip_writer.finish()?;
    Ok(())
}

/// Berechnet den SHA-512-Hash für ein gegebenes Passwort und einen Salt.
///
/// Ablauf:
/// 1. H0 = SHA512(Salt + Password(UTF-16LE))
/// 2. Loop `spin_count` times: H(n) = SHA512(H(n-1) + Iterator(4 Bytes LE))
fn compute_hash(password: &str, salt: &[u8], spin_count: u32) -> String {
    let pw_utf16: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&pw_utf16);
    let mut hash = hasher.finalize();

    for i in 0..spin_count {
        let mut iterator = [0u8; 4];
        let mut wtr = &mut iterator[..];
        wtr.write_u32::<LE>(i).unwrap();

        let mut next_hasher = Sha512::new();
        next_hasher.update(hash);
        next_hasher.update(iterator);
        hash = next_hasher.finalize();
    }

    general_purpose::STANDARD.encode(hash)
}

/// Erzeugt den Passwort-Hash mit zufälligem Salt (für Einzeldatei-Schutz).
fn hash_password(password: &str, spin_count: u32) -> (String, String) {
    let mut salt = [0u8; SALT_SIZE];
    getrandom::fill(&mut salt).expect("OS RNG unavailable");
    let hash = compute_hash(password, &salt, spin_count);
    (general_purpose::STANDARD.encode(salt), hash)
}

/// Fügt das <workbookProtection> Element in die workbook.xml ein.
///
/// Beachtet dabei die strikte XSD-Reihenfolge:
/// Das Element muss VOR <bookViews> oder <sheets> eingefügt werden.
fn inject_protection(xml_content: &[u8], salt: &str, hash: &str, spin_count: u32) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml_content);

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut inserted = false;

    let protection_elem_str = format!(
        r#"<workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="{}" workbookSaltValue="{}" workbookSpinCount="{}"/>"#,
        hash, salt, spin_count
    );

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name_str = std::str::from_utf8(e.name().into_inner())?;

                // Einfügepunkt finden: Vor sheets, bookViews oder functionGroups
                if !inserted
                    && (name_str == "sheets"
                        || name_str == "bookViews"
                        || name_str == "functionGroups")
                {
                    write_protection_tag(&mut writer, &protection_elem_str)?;
                    inserted = true;
                }

                // Falls bereits ein (ungeschütztes/anderes) workbookProtection Tag existiert -> ersetzen
                if name_str == "workbookProtection" {
                    let target_name = name_str.as_bytes().to_vec();
                    if !inserted {
                        write_protection_tag(&mut writer, &protection_elem_str)?;
                        inserted = true;
                    }

                    // Bestehendes Tag und dessen Inhalt überspringen
                    let mut depth = 1;
                    let mut skip_buf = Vec::new();
                    loop {
                        match reader.read_event_into(&mut skip_buf) {
                            Ok(Event::Start(e)) if e.name().into_inner() == target_name => {
                                depth += 1
                            }
                            Ok(Event::End(e)) if e.name().into_inner() == target_name => {
                                depth -= 1;
                                if depth == 0 {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                        skip_buf.clear();
                    }
                    continue;
                }
                writer.write_event(Event::Start(e.clone()))?;
            }
            Ok(Event::Empty(ref e)) => {
                let name_str = std::str::from_utf8(e.name().into_inner())?;

                // Einfügepunkt für Self-Closing Tags
                if !inserted && (name_str == "sheets" || name_str == "bookViews") {
                    write_protection_tag(&mut writer, &protection_elem_str)?;
                    inserted = true;
                }

                // Bestehendes workbookProtection Tag ersetzen
                if name_str == "workbookProtection" {
                    if !inserted {
                        write_protection_tag(&mut writer, &protection_elem_str)?;
                        inserted = true;
                    }
                    continue;
                }
                writer.write_event(Event::Empty(e.clone()))?;
            }
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(e.into()),
        }
        buf.clear();
    }
    Ok(writer.into_inner().into_inner())
}

/// Hilfsfunktion zum Schreiben des XML-Tags
fn write_protection_tag<W: std::io::Write>(writer: &mut Writer<W>, tag: &str) -> Result<()> {
    let mut temp_reader = Reader::from_str(tag);
    temp_reader.config_mut().trim_text_start = true;
    temp_reader.config_mut().trim_text_end = true;
    loop {
        match temp_reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(e) => writer.write_event(e)?,
            Err(e) => return Err(e.into()),
        }
    }
    Ok(())
}
