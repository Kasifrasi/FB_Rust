{
  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
    flake-utils.url = "github:numtide/flake-utils";
    devenv.url = "github:cachix/devenv";
    devenv.inputs.nixpkgs.follows = "nixpkgs";

    # --- DER FEHLENDE TEIL ---
    rust-overlay.url = "github:oxalica/rust-overlay";
    rust-overlay.inputs.nixpkgs.follows = "nixpkgs";
    # -------------------------

    dagger.url = "github:dagger/nix";
    dagger.inputs.nixpkgs.follows = "nixpkgs";
  };

  outputs = { self, nixpkgs, flake-utils, devenv, ... } @ inputs:
    flake-utils.lib.eachDefaultSystem (system:
      let
        pkgs = import nixpkgs {
          inherit system;
          config.allowUnfree = true;
          # Optional: Falls du das Overlay direkt in pkgs nutzen willst
          overlays = [ (import inputs.rust-overlay) ];
        };
      in {
        formatter = pkgs.nixpkgs-fmt;

        devShells = {
          default = devenv.lib.mkShell {
            inherit pkgs inputs; # Hier werden die inputs (inkl. rust-overlay) an devenv gereicht
            modules = [
              ({ ... }: { _module.args.self = self; })
              ./devenv.nix
            ];
          };

          testing = devenv.lib.mkShell {
            inherit pkgs inputs;
            modules = [
              ({ ... }: { _module.args.self = self; })
              ./nix/gui-test.nix
            ];
          };
        };
      });
}
