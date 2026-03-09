# prek.nix
{ pkgs ? import <nixpkgs> { } }:
pkgs.rustPlatform.buildRustPackage {
  pname = "prek";
  version = "latest-git";
  src = pkgs.fetchFromGitHub {
    owner = "j178";
    repo = "prek";
    rev = "master"; 
    hash = ""; # Hier den Hash einfügen oder beim ersten Run von Nix ausgeben lassen
  };
  cargoLock.lockFile = null; # Falls kein Lock im Repo ist, sonst Pfad angeben
  doCheck = false;
}