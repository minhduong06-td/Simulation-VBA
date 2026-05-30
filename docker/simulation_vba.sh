#!/usr/bin/env bash
set -Eeuo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd -P)"
CALLER_DIR="$(pwd -P)"
IMAGE="${SIMULATIONVBA_DOCKER_IMAGE:-haroldogden/vipermonkey:latest}"
PULL_IMAGE="${SIMULATIONVBA_DOCKER_PULL:-1}"
KEEP_CONTAINER="${SIMULATIONVBA_DOCKER_KEEP:-0}"

usage() {
    cat <<'USAGE'
Usage: simulation_vba.sh FILE [JSON_FILE] [-i ENTRY]
       simulation_vba.sh FILE -i ENTRY

Runs each analysis in a brand-new disposable Docker container.
The local SimulationVBA source tree is copied into the container every run.

Environment variables:
  SIMULATIONVBA_DOCKER_IMAGE  Docker image to use as base runtime
                              default: haroldogden/vipermonkey:latest
  SIMULATIONVBA_DOCKER_PULL   Pull image before each run: 1=yes, 0=no
                              default: 1
  SIMULATIONVBA_DOCKER_KEEP   Keep container after run for debugging: 1=yes, 0=no
                              default: 0
USAGE
}

if [[ $# -eq 0 || "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
    usage
    exit 0
fi

input_file="$1"
if [[ ! -f "$input_file" ]]; then
    echo "[!] Input file not found: $input_file" >&2
    exit 1
fi

if [ "$(uname)" == "Darwin" ]; then
    echo "[*] User running on a Mac"
    if command -v docker-machine >/dev/null 2>&1 && [ "$(docker-machine status 2>/dev/null || true)" == "Stopped" ]; then
        echo "[*] 'docker-machine' is stopped. Starting it and instantiating the environment."
        docker-machine start
        eval "$(docker-machine env)"
    fi
fi

echo "[*] Running 'docker ps' to see if script has required privileges to run..."
if ! docker ps >/dev/null; then
    echo "[!] 'docker ps' failed - you may not have privileges to run docker. Try sudo or join the docker group." >&2
    exit 1
fi

if [[ "$PULL_IMAGE" != "0" ]]; then
    echo "[*] Pulling Docker image $IMAGE..."
    docker pull "$IMAGE"
fi

run_id="$(date +%Y%m%d%H%M%S)-$$-$RANDOM"
container_name="simulation-vba-run-$run_id"
docker_id=""
file_basename="$(basename "$input_file")"
container_input="/root/input/$file_basename"
container_json="/root/output/report.json"
container_artifact_dir="/root/input/${file_basename}_artifacts"
container_zip="/root/output/${file_basename}_artifacts.zip"

cleanup() {
    status=$?
    if [[ -n "${docker_id:-}" ]]; then
        if [[ "$KEEP_CONTAINER" == "1" ]]; then
            echo "[*] Keeping container for debugging: $docker_id"
        else
            echo "[*] Removing docker container $docker_id"
            docker rm -f "$docker_id" >/dev/null 2>&1 || true
        fi
    fi
    exit "$status"
}
trap cleanup EXIT INT TERM

echo "[*] Starting fresh isolated container..."
docker_id="$(docker run -d -t --network none --name "$container_name" "$IMAGE")"

echo "[*] Container ID: $docker_id"
echo "[*] Preparing clean runtime directories..."
docker exec "$docker_id" sh -c 'rm -rf /opt/simulation_vba /root/input /root/output /root/.cache; mkdir -p /opt/simulation_vba /root/input /root/output'

echo "[*] Copying local SimulationVBA source into container..."
tar \
    --exclude='.git' \
    --exclude='__pycache__' \
    --exclude='*.pyc' \
    --exclude='*.pyo' \
    --exclude='*.egg-info' \
    --exclude='build' \
    --exclude='dist' \
    --exclude='test.zip' \
    --exclude='*_artifacts.zip' \
    -C "$REPO_ROOT" -cf - . | docker exec -i "$docker_id" tar -xf - -C /opt/simulation_vba

docker exec "$docker_id" sh -c 'chmod +x /opt/simulation_vba/simulation_vba/vba_emu.py 2>/dev/null || true'

echo "[*] Selecting Python runtime with SimulationVBA dependencies..."
vm_python="$(docker exec "$docker_id" sh -c '
    cd /opt/simulation_vba || exit 1
    for py in pypy pypy3 python2 python3 python; do
        if command -v "$py" >/dev/null 2>&1; then
            if "$py" -c "import colorlog, prettytable, oletools, olefile, pyparsing" >/dev/null 2>&1; then
                printf "%s" "$py"
                exit 0
            fi
        fi
    done
    exit 42
' 2>/dev/null || true)"

if [[ -z "$vm_python" ]]; then
    echo "[!] Could not find a Python/PyPy runtime in the container with required SimulationVBA dependencies." >&2
    echo "[!] The base image is expected to provide colorlog, prettytable, oletools, olefile, and pyparsing." >&2
    echo "[!] Try rebuilding the Docker image with requirements.txt, or use a base image that already contains SimulationVBA dependencies." >&2
    exit 1
fi
echo "[*] Using runtime: $vm_python"

echo "[*] Copying sample into container..."
docker cp "$input_file" "$docker_id:$container_input"

echo "[*] Starting openoffice listener for file content conversions..."
docker exec "$docker_id" sh -c '/usr/lib/libreoffice/program/soffice.bin --headless --invisible --nocrashreport --nodefault --nofirststartwizard --nologo --norestore --accept="socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext" >/tmp/soffice.log 2>&1 &'

entry=""
json=""
json_file=""

if [[ $# -ge 3 && "${2:-}" == "-i" ]]; then
    entry="-i $3"
elif [[ $# -eq 2 ]]; then
    json="-o $container_json"
    json_file="$2"
fi

if [[ $# -ge 4 && "${3:-}" == "-i" ]]; then
    entry="-i $4"
    json="-o $container_json"
    json_file="$2"
fi

json_out=""
if [[ -n "$json_file" ]]; then
    if [[ "$json_file" = /* ]]; then
        json_out="$json_file"
    else
        json_out="$CALLER_DIR/$json_file"
    fi
    rm -f "$json_out"
fi

out_zip="$CALLER_DIR/${file_basename}_artifacts.zip"
out_dir="$CALLER_DIR/${file_basename}_artifacts"
rm -rf "$out_dir" "$out_zip"

tmp_zip="$(mktemp "${TMPDIR:-/tmp}/simulation_vba-artifacts.XXXXXX.zip")"

echo "[*] Running SimulationVBA from copied local source..."
docker exec "$docker_id" sh -c "cd /opt/simulation_vba && $vm_python simulation_vba/vba_emu.py -s --ioc --jit '$container_input' $json $entry"

if [[ -n "$json_file" ]]; then
    if docker exec "$docker_id" test -s "$container_json"; then
        docker cp "$docker_id:$container_json" "$json_out"
        echo "[*] JSON report saved to $json_out"
    else
        echo "[*] No JSON report was generated."
    fi
fi

has_artifacts=false
if docker exec "$docker_id" sh -c "ls '$container_artifact_dir/' 2>/dev/null | head -n 1 | grep -q ."; then
    has_artifacts=true
    echo "[*] Copying dropped files from this run to host..."
    docker cp "$docker_id:$container_artifact_dir/." "$out_dir"
    docker exec "$docker_id" sh -c "cd /root/input && zip -q -r --password=infected '$container_zip' '${file_basename}_artifacts'"
    docker cp "$docker_id:$container_zip" "$tmp_zip"
    if [[ -s "$tmp_zip" ]]; then
        mv -f "$tmp_zip" "$out_zip"
    fi
fi

if $has_artifacts; then
    echo ""
    echo "============================================"
    echo "Artifacts copied to:"
    echo "  $out_dir"
    echo "  $out_zip"
    echo "============================================"
else
    rm -f "$tmp_zip"
    echo "[*] No dropped files were produced in this run."
fi

echo "[*] Done."
