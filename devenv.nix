{ pkgs, lib, ... }:

{
  # 1. Wir nutzen das eingebaute Rust-Modul von devenv.
  # Das kümmert sich automatisch um Environment-Variablen wie RUST_SRC_PATH.
  languages.rust = {
    enable = true;
    channel = "stable"; # Wechsel hier auf "nightly" falls nötig
    
    # Diese Komponenten sind PFLICHT für eine gute IDE-Erfahrung:
    components = [ 
      "rustc" 
      "cargo" 
      "clippy" 
      "rustfmt" 
      "rust-analyzer" 
      "rust-src"  # <--- GANZ WICHTIG: Ohne das geht "Go to Definition" für std-lib nicht!
    ];
  };

  packages = [
    # Hilfstools, damit Rust C-Bibliotheken findet
    pkgs.pkg-config 
    pkgs.openssl
    pkgs.lldb
  ];

  env = {
    # 2. Der wichtigste Teil gegen Linker-Fehler (wie "shared object not found")!
    # Rust-Binaries brauchen Zugriff auf System-Bibliotheken.
    LD_LIBRARY_PATH = lib.makeLibraryPath [
      pkgs.stdenv.cc.cc.lib  # libstdc++ etc.
      pkgs.openssl           # Für fast alles, was Netzwerk macht
      # pkgs.zlib            # Falls du zlib brauchst
      # pkgs.wayland         # Falls du GUI Apps baust
    ];
    
    # Erzwingt, dass pkg-config auch openssl findet
    PKG_CONFIG_PATH = "${pkgs.openssl.dev}/lib/pkgconfig";
  };

  enterShell = ''
    echo "🦀 Rust Umgebung geladen"
    
    # Check, ob alles da ist
    echo "✅ Compiler: $(rustc --version)"
    echo "✅ Analyzer: $(rust-analyzer --version)"
    
    # Kleiner Check, ob die Source-Pfade stimmen (für Debugging)
    if [ -z "$RUST_SRC_PATH" ]; then
       echo "⚠️ Warnung: RUST_SRC_PATH ist nicht gesetzt!"
    else
       echo "✅ Sources gefunden für VSCode"
    fi
  '';
}
