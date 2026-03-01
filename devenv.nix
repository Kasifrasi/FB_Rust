{ pkgs, ... }:
{
  languages.rust = {
    enable = true;
    channel = "stable";
    mold.enable = false; # Wir machen es manuell
    components = [ "rustc" "cargo" "clippy" "rustfmt" "rust-analyzer" ];
  };

  env = {
    CC = "clang";
    CXX = "clang++";
    RUSTC_WRAPPER = "sccache";
    CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_LINKER = "clang";
    CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS = "-C link-arg=-fuse-ld=mold";
  };

  # Projekt-Tools im Root
  packages = with pkgs; [
      prek
      cargo-expand
      cargo-license
      cargo-about
      cargo-llvm-cov
      cargo-audit
      cargo-nextest
      cargo-deny
      cargo-cyclonedx
      cargo-edit
      
      pkg-config 
      mold
      clang 
      sccache 
      bacon 
      cargo-nextest
    ];
    
  # Einfacher Start-Check
  enterShell = ''
      if [ ! -d ".venv" ]; then uv venv; fi
      source .venv/bin/activate
      
      # Prek Hooks installieren/aktualisieren
      if [ -d ".git" ]; then
        prek install
      fi
      
      echo "🦀 Rust Dev | 🛡️ prek Hooks"
    '';
}