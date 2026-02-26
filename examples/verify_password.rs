//! Verify Workbook Protection Password Hash
//!
//! This tool helps verify that the password hashing is working correctly

use base64::{engine::general_purpose, Engine as _};
use byteorder::{WriteBytesExt, LE};
use sha2::{Digest, Sha512};

const SPIN_COUNT: u32 = 100_000;

fn hash_password_with_salt(password: &str, salt: &[u8]) -> String {
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

    general_purpose::STANDARD.encode(hash)
}

fn main() {
    println!("Workbook Protection Password Hash Verifier\n");

    // Read the salt from the generated file
    let output_file = "tests/output/test_with_wb_protection.xlsx";

    println!("Reading workbook.xml from: {}", output_file);

    // Try to extract salt and hash from XML
    let result = std::process::Command::new("unzip")
        .args(["-p", output_file, "xl/workbook.xml"])
        .output();

    if let Ok(output) = result {
        let xml = String::from_utf8_lossy(&output.stdout);

        // Extract salt and hash
        if let Some(salt_start) = xml.find("workbookSaltValue=\"") {
            let salt_start = salt_start + 19;
            if let Some(salt_end) = xml[salt_start..].find('"') {
                let salt_b64 = &xml[salt_start..salt_start + salt_end];

                if let Some(hash_start) = xml.find("workbookHashValue=\"") {
                    let hash_start = hash_start + 19;
                    if let Some(hash_end) = xml[hash_start..].find('"') {
                        let stored_hash_b64 = &xml[hash_start..hash_start + hash_end];

                        println!("Salt (Base64): {}", salt_b64);
                        println!("Stored Hash (Base64): {}", stored_hash_b64);

                        // Decode salt
                        if let Ok(salt) = general_purpose::STANDARD.decode(salt_b64) {
                            // Test with the password
                            let test_password = "secret123";
                            let computed_hash = hash_password_with_salt(test_password, &salt);

                            println!("\nTesting password: '{}'", test_password);
                            println!("Computed Hash (Base64): {}", computed_hash);

                            if computed_hash == stored_hash_b64 {
                                println!("\n✅ PASSWORD MATCHES! Hash is correct.");
                            } else {
                                println!("\n❌ PASSWORD DOES NOT MATCH!");
                                println!("   This could mean:");
                                println!("   1. The password is incorrect");
                                println!("   2. The hashing algorithm has a bug");
                                println!("   3. The Excel implementation differs");
                            }
                        }
                    }
                }
            }
        }
    } else {
        println!("Error: Could not read the workbook file");
        println!("Please ensure the file exists and unzip is installed");
    }
}
