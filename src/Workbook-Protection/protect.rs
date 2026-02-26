use anyhow::{Context, Result};
use base64::{engine::general_purpose, Engine as _};
use byteorder::{WriteBytesExt, LE};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use rand::RngCore;
use sha2::{Digest, Sha512};
use std::fs::File;
use std::io::{Cursor, Read, Write};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const SPIN_COUNT: u32 = 100_000;
const SALT_SIZE: usize = 16;

/// Schützt die Arbeitsmappenstruktur einer existierenden Excel-Datei mit einem Passwort.
///
/// Dies verhindert, dass Benutzer Blätter hinzufügen, löschen, verschieben oder umbenennen.
/// Die Implementierung folgt dem ECMA-376 Standard (Agile Encryption mit SHA-512).
///
/// # Argumente
/// * `input_path` - Pfad zur ungeschützten .xlsx Datei
/// * `output_path` - Pfad, wo die geschützte Datei gespeichert werden soll
/// * `password` - Das zu setzende Passwort
pub fn protect_workbook(input_path: &str, output_path: &str, password: &str) -> Result<()> {
    let file = File::open(input_path).context("Konnte Input-Datei nicht öffnen")?;
    let mut archive = ZipArchive::new(file).context("Konnte ZIP-Archiv nicht lesen")?;

    let out_file = File::create(output_path).context("Konnte Output-Datei nicht erstellen")?;
    let mut zip_writer = ZipWriter::new(out_file);

    // Passwort-Hash berechnen
    let (salt_b64, hash_b64) = hash_password(password);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();

        let mut content = Vec::new();
        file.read_to_end(&mut content)?;

        let options = FileOptions::default()
            .compression_method(file.compression())
            .unix_permissions(file.unix_mode().unwrap_or(0o644));

        zip_writer.start_file(&name, options)?;

        if name == "xl/workbook.xml" {
            // workbook.xml manipulieren, um workbookProtection einzufügen
            let new_xml = inject_protection(&content, &salt_b64, &hash_b64)?;
            zip_writer.write_all(&new_xml)?;
        } else {
            zip_writer.write_all(&content)?;
        }
    }

    zip_writer.finish()?;
    Ok(())
}

/// Erzeugt den Passwort-Hash gemäß Excel 2013+ Standard (Agile Encryption).
///
/// Ablauf:
/// 1. Salt generieren (16 Bytes)
/// 2. H0 = SHA512(Salt + Password(UTF-16LE))
/// 3. Loop 100.000x: H(n) = SHA512(H(n-1) + Iterator(4 Bytes LE))
fn hash_password(password: &str) -> (String, String) {
    let mut rng = rand::thread_rng();
    let mut salt = [0u8; SALT_SIZE];
    rng.fill_bytes(&mut salt);

    let pw_utf16: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&pw_utf16);
    let mut hash = hasher.finalize();

    for i in 0..SPIN_COUNT {
        let mut iterator = [0u8; 4];
        let mut wtr = &mut iterator[..];
        wtr.write_u32::<LE>(i).unwrap();

        let mut next_hasher = Sha512::new();
        next_hasher.update(hash);
        next_hasher.update(iterator);
        hash = next_hasher.finalize();
    }

    (
        general_purpose::STANDARD.encode(salt),
        general_purpose::STANDARD.encode(hash),
    )
}

/// Fügt das <workbookProtection> Element in die workbook.xml ein.
///
/// Beachtet dabei die strikte XSD-Reihenfolge:
/// Das Element muss VOR <bookViews> oder <sheets> eingefügt werden.
fn inject_protection(xml_content: &[u8], salt: &str, hash: &str) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml_content);
    reader.trim_text(false);

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut inserted = false;

    let protection_elem_str = format!(
        r#"<workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="{}" workbookSaltValue="{}" workbookSpinCount="{}"/>"#,
        hash, salt, SPIN_COUNT
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
    temp_reader.trim_text(true);
    while let Ok(e) = temp_reader.read_event() {
        if let Event::Eof = e {
            break;
        }
        writer.write_event(e)?;
    }
    Ok(())
}
