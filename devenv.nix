{ pkgs, ... }:
{
  # Hier werden alle Module aus dem nix-Ordner geladen
  imports = [
    ./nix/gui-support.nix
    ./nix/toolchain.nix
    ./nix/dagger.nix # Falls du dagger.nix auch dorthin schiebst
  ];

  # Projekt-Tools im Root
  packages = [
      pkgs.prek
      pkgs.xdg-utils
      pkgs.flatpak
      pkgs.flatpak-builder
      pkgs.cargo-expand
      pkgs.cargo-license
      pkgs.cargo-about
      pkgs.cargo-llvm-cov
      pkgs.cargo-audit
      pkgs.cargo-nextest
      pkgs.cargo-deny
      pkgs.cargo-cyclonedx
      pkgs.python3Packages.pyperf   # System-Tuning für reproduzierbare Benchmarks
    ];
    
  # Einfacher Start-Check
  enterShell = ''
      if [ ! -d ".venv" ]; then uv venv; fi
      source .venv/bin/activate
      
      # Prek Hooks installieren/aktualisieren
      if [ -d ".git" ]; then
        prek install
      fi
      
      echo "🐍 Python Dev | 🦀 Rust Dev | 🎨 GUI Support | 🛡️ prek Hooks"
      echo "Status GSettings: $([ -f "$GSETTINGS_SCHEMA_DIR/gschemas.compiled" ] && echo "OK ✅" || echo "Fehler ❌")"
    '';
}