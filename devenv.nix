{ pkgs, lib, ... }:
{
  # Hier werden alle Module aus dem nix-Ordner geladen
  imports = [
    ./nix/gui-support.nix
    ./nix/toolchain.nix
    ./nix/dagger.nix # Falls du dagger.nix auch dorthin schiebst
  ];

  # Projekt-Tools im Root
  packages = [ pkgs.xdg-utils pkgs.flatpak pkgs.flatpak-builder ];

  # Einfacher Start-Check
  enterShell = ''
    if [ ! -d ".venv" ]; then uv venv; fi
    source .venv/bin/activate
    
    echo "🐍 Python Dev | 🦀 Rust Dev | 🎨 GUI Support"
    echo "Status GSettings: $([ -f "$GSETTINGS_SCHEMA_DIR/gschemas.compiled" ] && echo "OK ✅" || echo "Fehler ❌")"
  '';
}
