#!/usr/bin/env nu

# --- Konfiguration ---
let app_id = "com.ardit.tauriapp"
let binary_name = "tauri-app"
let runtime = "org.gnome.Platform"
let sdk = "org.gnome.Sdk"
let version = "47"
let output_file = $"($app_id).flatpak"

def main [] {
    print $"🚀 Starte Packaging für ($app_id)..."

    # 1. Checks
    if not ("dist-linux" | path exists) {
        error make {msg: "❌ Ordner ./dist-linux fehlt! Bitte erst 'dagger call' ausführen."}
    }

    # 2. Aufräumen
    print "🧹 Bereinige alte Builds..."
    # Wir löschen alles, was im Weg ist
    try { ^rm -rf flatpak-build repo } 
    sleep 100ms 
    
    # WICHTIG: Hier KEIN mkdir flatpak-build machen! 
    # flatpak build-init will den Ordner selbst erstellen.

    # 3. Init
    print "⚙️  Initialisiere Flatpak Build-Dir..."
    # Dieser Befehl erstellt den Ordner 'flatpak-build' und die Unterordner
    flatpak build-init flatpak-build $app_id $sdk $runtime $version

    # 4. Dateien kopieren
    print "📂 Kopiere Dateien..."
    # Jetzt, wo build-init fertig ist, existiert auch flatpak-build/files
    cp -r dist-linux/* flatpak-build/files/

    # 5. Finish
    print "🔒 Setze Permissions..."
    let permissions = [
        "--share=ipc"
        "--socket=x11"
        "--socket=wayland"
        "--device=dri"
        "--share=network"
        $"--command=($binary_name)"
    ]
    flatpak build-finish flatpak-build ...$permissions

    # 6. Export & Bundle
    print "📦 Erstelle Bundle..."
    mkdir repo
    flatpak build-export repo flatpak-build
    flatpak build-bundle repo $output_file $app_id

    print $"✅ Erfolg: ($output_file)"
}
