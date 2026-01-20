{ pkgs, lib, ... }:
let
  compiledSchemas = pkgs.runCommand "compiled-gsettings-schemas"
    { nativeBuildInputs = [ pkgs.glib pkgs.gtk3 ]; } ''
    schemaDir=$out/share/glib-2.0/schemas
    mkdir -p $schemaDir
    for pkg in ${pkgs.gsettings-desktop-schemas} ${pkgs.gtk3} ${pkgs.glib-networking}; do
      [ -d "$pkg/share/glib-2.0/schemas" ] && cp -rf "$pkg/share/glib-2.0/schemas/"*.xml $schemaDir/ 2>/dev/null || true
    done
    glib-compile-schemas $schemaDir
  '';
in {
  packages = with pkgs; [
    adwaita-icon-theme hicolor-icon-theme librsvg
    glib gtk3 webkitgtk_4_1
    xdg-desktop-portal xdg-desktop-portal-gtk
  ];

  env = {
    GSETTINGS_SCHEMA_DIR = "${compiledSchemas}/share/glib-2.0/schemas";
    XDG_DATA_DIRS = lib.makeSearchPath "share" [
      compiledSchemas
      pkgs.adwaita-icon-theme
      pkgs.hicolor-icon-theme
      pkgs.gtk3
    ];
    GIO_EXTRA_MODULES = "${pkgs.glib-networking}/lib/gio/modules";
    XDG_CURRENT_DESKTOP = "GNOME";
  };
}
