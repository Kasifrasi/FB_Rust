{ pkgs, ... }:
{
  languages.rust = {
    enable = true;
    channel = "stable";
    mold.enable = false; # Wir machen es manuell
    components = [ "rustc" "cargo" "clippy" "rustfmt" "rust-analyzer" ];
  };

  packages = with pkgs; [
    python314 uv nodejs_24 pnpm pkg-config mold
    clang sccache bacon cargo-nextest
  ];

  env = {
    CC = "clang";
    CXX = "clang++";
    RUSTC_WRAPPER = "sccache";
    CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_LINKER = "clang";
    CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS = "-C link-arg=-fuse-ld=mold";
  };
}
