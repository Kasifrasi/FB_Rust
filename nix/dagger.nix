{ pkgs, inputs, lib, ... }:

{
  packages = [
    inputs.dagger.packages.${pkgs.stdenv.hostPlatform.system}.dagger
  ];

  env = {
    DOCKER_HOST = "unix:///run/user/$(id -u)/podman/podman.sock";
    _EXPERIMENTAL_DAGGER_RUNNER_HOST = "docker-container://dagger-engine";
  };

  enterShell = ''
    # --- DAGGER ENGINE START LOGIK ---
    CONTAINER_NAME="dagger-engine"
    DAGGER_CLI_VER=$(dagger version | head -n1 | awk '{print $2}')
    if [ -z "$DAGGER_CLI_VER" ]; then DAGGER_CLI_VER="v0.19.8"; fi
    ENGINE_IMAGE="registry.dagger.io/engine:$DAGGER_CLI_VER"

    if ! podman ps --format "{{.Names}}" | grep -q "^$CONTAINER_NAME$"; then
        echo "⚙️  Starte Dagger Engine ($DAGGER_CLI_VER)..."
        podman rm -f $CONTAINER_NAME >/dev/null 2>&1 || true
        podman run -d \
            --name $CONTAINER_NAME \
            --privileged \
            --security-opt label=disable \
            --restart always \
            -v dagger-engine-data:/var/lib/dagger \
            -v "/run/user/$(id -u)/podman/podman.sock:/var/run/docker.sock" \
            "$ENGINE_IMAGE" > /dev/null
        sleep 3
    else
        echo "⚡ Dagger Engine läuft bereits."
    fi
  '';
}
