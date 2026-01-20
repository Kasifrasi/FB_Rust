{ pkgs, ... }:

let
  compiledSchemas = pkgs.runCommand "compiled-gsettings-schemas" { nativeBuildInputs = [ pkgs.glib ]; } ''
    mkdir -p $out/share/glib-2.0/schemas
    cp -rf "${pkgs.gsettings-desktop-schemas}/share/gsettings-schemas/${pkgs.gsettings-desktop-schemas.name}/glib-2.0/schemas/"* $out/share/glib-2.0/schemas/
    cp -rf "${pkgs.gtk3}/share/gsettings-schemas/${pkgs.gtk3.name}/glib-2.0/schemas/"* $out/share/glib-2.0/schemas/
    glib-compile-schemas $out/share/glib-2.0/schemas
  '';
in {
  # Wir bauen auf der normalen Config auf
  imports = [ ../devenv.nix ];

  packages = [ compiledSchemas ];

  env.XDG_DATA_DIRS = "${compiledSchemas}/share:$XDG_DATA_DIRS";

  enterShell = ''
    echo "🎨 GUI-Test-Umgebung geladen. Schemas sind aktiv."
  '';
}
