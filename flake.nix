{
  inputs = {
    # 1. Stabiler Channel für minimale Downloads und einen warmen Cache
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-25.11";
    flake-utils.url = "github:numtide/flake-utils";
    devenv.url = "github:cachix/devenv";
    
    # 2. Prek direkt vom Git-Master ohne Flake-Support
    prek.url = "github:j178/prek";
    prek.flake = false; 

    rust-overlay.url = "github:oxalica/rust-overlay";
    rust-overlay.inputs.nixpkgs.follows = "nixpkgs";
    
    dagger.url = "github:dagger/nix";
    dagger.inputs.nixpkgs.follows = "nixpkgs";
  };

  outputs = { self, nixpkgs, flake-utils, devenv, prek, ... } @ inputs:
      flake-utils.lib.eachDefaultSystem (system:
        let
          pkgs = import nixpkgs {
            inherit system;
            config.allowUnfree = true;
            overlays = [ 
              (import inputs.rust-overlay) 
              
              # NEU: Das Overlay! Es überschreibt das Standard-Prek im gesamten System
              (final: prev: {
                prek = prev.rustPlatform.buildRustPackage {
                  pname = "prek";
                  version = "latest-git";
                  src = prek;
                  cargoLock.lockFile = "${prek}/Cargo.lock";
                  doCheck = false; # Tests überspringen
                };
              })
            ];
          };
  
        in {
          formatter = pkgs.nixpkgs-fmt;
  
          devShells = {
            default = devenv.lib.mkShell {
              inherit pkgs inputs; 
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