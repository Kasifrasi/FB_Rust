{
  description = "Prek tool - standalone build";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
  };

  outputs = { self, nixpkgs }:
    let
      # Unterstützung für Linux und macOS
      systems = [ "x86_64-linux" "aarch64-linux" "x86_64-darwin" "aarch64-darwin" ];
      forAllSystems = f: nixpkgs.lib.genAttrs systems (system: f nixpkgs.legacyPackages.${system});
    in
    {
      # Hier definieren wir das Paket direkt
      packages = forAllSystems (pkgs: {
        default = pkgs.rustPlatform.buildRustPackage {
          pname = "prek";
          version = "0.3.5";

          src = pkgs.fetchFromGitHub {
            owner = "j178";
            repo = "prek";
            rev = "v0.3.5"; # Wir nutzen den Tag passend zur Version
            # DEINE HASHES
            hash = "sha256-XWUotVd6DGk8IfE5UT2NjgSB6FL/HDEBr/wBFqOMe0I=";
          };

          cargoHash = "sha256-ZIkbA6rfS+8YhfP0YE4v9Me9FeRvLVLGRBUZnoA9ids=";

          nativeBuildInputs = [ pkgs.pkg-config ];

          # Wir deaktivieren die Tests, da sie oft Netzwerkzugriff brauchen
          doCheck = false;

          meta = {
            description = "Better pre-commit, re-engineered in Rust";
            homepage = "https://github.com/j178/prek";
            license = pkgs.lib.licenses.mit;
            mainProgram = "prek";
          };
        };
      });
    };
}