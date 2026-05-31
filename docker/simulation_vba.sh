#!/usr/bin/env bash
set -Eeuo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd -P)"
CALLER_DIR="$(pwd -P)"
IMAGE="${SIMULATIONVBA_DOCKER_IMAGE:-md06/simulation:latest}"
PULL_IMAGE="${SIMULATIONVBA_DOCKER_PULL:-0}"
KEEP_CONTAINER="${SIMULATIONVBA_DOCKER_KEEP:-0}"

REBUILD_IMAGE=0
CLEAN_DOCKER=0
CLEAN_ARTIFACTS=0

usage() {
    cat <<'USAGE'
Usage: simulation_vba.sh [OPTIONS] FILE [VBA_OPTIONS...]
       simulation_vba.sh --clean
       simulation_vba.sh --clean-artifacts
       simulation_vba.sh --rebuild [FILE [VBA_OPTIONS...]]
       simulation_vba.sh [VBA_OPTIONS...] FILE

Runs each analysis in a brand-new disposable Docker container.
The local SimulationVBA source tree is copied into the container every run.

New options (processed before FILE):
  --clean           Remove all simulation Docker containers, temp images,
                    and the cached image (md06/simulation:latest).
                    Also remove repo build artifacts (*_artifacts, *.pyc, ...).
                    Does NOT remove source code.
  --clean-artifacts Remove only repo build artifacts (*_artifacts, *.pyc,
                    __pycache__, .pytest_cache). Does NOT touch Docker.
  --rebuild         Force rebuild the cached image from docker/Dockerfile,
                    then run FILE if provided.
  --help, -h        Show this help.

VBA_OPTIONS are forwarded verbatim to vba_emu.py (e.g. --deob simulate).
They can appear before or after FILE.

Environment variables:
  SIMULATIONVBA_DOCKER_IMAGE   Docker image tag to use (default: md06/simulation:latest)
  SIMULATIONVBA_DOCKER_PULL    Pull image before each run: 1=yes, 0=no (default: 0)
  SIMULATIONVBA_DOCKER_KEEP    Keep container after run for debugging: 1=yes, 0=no (default: 0)
USAGE
}

# ----- helper functions -----

clean_artifacts() {
    find "$REPO_ROOT" -type d -name "__pycache__" -prune -exec rm -rf {} + 2>/dev/null || true
    find "$REPO_ROOT" -name "*.pyc" -delete 2>/dev/null || true
    rm -rf "$REPO_ROOT"/.pytest_cache 2>/dev/null || true
    rm -rf "$REPO_ROOT"/*_artifacts "$REPO_ROOT"/*_artifacts.zip 2>/dev/null || true
}

clean_docker() {
    local ids
    ids="$(docker ps -a --filter "name=simulation-vba-run-" -q 2>/dev/null)"
    if [[ -n "$ids" ]]; then
        echo "[*] Removing simulation-vba-run-* containers..."
        echo "$ids" | xargs -r docker rm -f >/dev/null 2>&1 || true
    fi
    local temp_images
    temp_images="$(docker images --format "{{.Repository}}:{{.Tag}}" 2>/dev/null | grep "^simulation-vba-temp:" || true)"
    if [[ -n "$temp_images" ]]; then
        echo "[*] Removing simulation-vba-temp:* images..."
        echo "$temp_images" | xargs -r docker rmi -f >/dev/null 2>&1 || true
    fi
    if docker image inspect "$IMAGE" >/dev/null 2>&1; then
        echo "[*] Removing cached image $IMAGE..."
        docker rmi -f "$IMAGE" >/dev/null 2>&1 || true
    fi
}

build_image() {
    echo "[*] Building Docker image $IMAGE from docker/Dockerfile..."
    clean_artifacts
    docker build -t "$IMAGE" -f "$REPO_ROOT/docker/Dockerfile" "$REPO_ROOT"
}

ensure_image() {
    if docker image inspect "$IMAGE" >/dev/null 2>&1; then
        echo "[*] Using cached image: $IMAGE"
    else
        echo "[*] Image $IMAGE not found locally."
        build_image
    fi
    if [[ "$PULL_IMAGE" != "0" ]]; then
        echo "[*] Pulling Docker image $IMAGE..."
        docker pull "$IMAGE"
    fi
}

# ----- parse global options -----

global_args=()
while [[ $# -gt 0 ]]; do
    case "$1" in
        --clean)
            CLEAN_DOCKER=1
            shift
            ;;
        --clean-artifacts)
            CLEAN_ARTIFACTS=1
            shift
            ;;
        --rebuild)
            REBUILD_IMAGE=1
            shift
            ;;
        --help|-h)
            usage
            exit 0
            ;;
        --deob)
            # Consume --deob and its value; validate mode
            if [[ $# -lt 2 ]]; then
                echo "[!] --deob requires an argument (simulate)" >&2
                exit 1
            fi
            deob_mode="$2"
            shift 2
            if [[ "$deob_mode" != "simulate" ]]; then
                echo "[!] Unsupported --deob mode: $deob_mode. Supported: simulate" >&2
                exit 1
            fi
            global_args+=("--deob" "simulate")
            ;;
        --)
            shift
            global_args+=("$@")
            break
            ;;
        *)
            global_args+=("$1")
            shift
            ;;
    esac
done

set -- "${global_args[@]}"

if [[ "$CLEAN_DOCKER" == "1" ]]; then
    clean_docker
    clean_artifacts
    echo "[*] Clean complete."
    exit 0
fi

if [[ "$CLEAN_ARTIFACTS" == "1" ]]; then
    clean_artifacts
    echo "[*] Artifact clean complete."
    exit 0
fi

if [[ "$REBUILD_IMAGE" == "1" ]]; then
    docker rmi -f "$IMAGE" 2>/dev/null || true
    build_image
fi

# After global options, determine input file and forward remaining args to vba_emu
if [[ $# -eq 0 ]]; then
    usage
    exit 1
fi

# Option-first: if first arg starts with -, treat last arg as the file
if [[ "$#" -gt 1 && "${1:-}" == -* ]]; then
    input_candidate="${@: -1}"
    leading_opts=("${@:1:$#-1}")
    set -- "$input_candidate" "${leading_opts[@]}"
fi

input_file="$1"
if [[ ! -f "$input_file" ]]; then
    echo "[!] Input file not found: $input_file" >&2
    exit 1
fi

# ----- Docker capability check -----

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

# ----- ensure image is available -----

ensure_image

# ----- prepare runtime -----

run_id="$(date +%Y%m%d%H%M%S)-$$-$RANDOM"
container_name="simulation-vba-run-$run_id"
docker_id=""
file_basename="$(basename "$input_file")"
container_input="/root/input/$file_basename"
container_json="/root/output/report.json"
container_artifact_dir="/root/input/${file_basename}_artifacts"
container_zip="/root/output/${file_basename}_artifacts.zip"

cleanup() {
    local status=$?
    if [[ -n "${docker_id:-}" ]]; then
        if [[ "$KEEP_CONTAINER" == "1" ]]; then
            echo "[*] Keeping container for debugging: $docker_id"
        else
            docker rm -f "$docker_id" >/dev/null 2>&1 || true
        fi
    fi
    exit "$status"
}
trap cleanup EXIT INT TERM HUP

echo "[*] Starting fresh isolated container..."
docker_id="$(docker run -d -t --network none --name "$container_name" "$IMAGE")"

echo "[*] Container ID: $docker_id"
echo "[*] Preparing runtime directories..."
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
    --exclude='*_artifacts' \
    --exclude='*_artifacts.zip' \
    -C "$REPO_ROOT" -cf - . | docker exec -i "$docker_id" tar -xf - -C /opt/simulation_vba

docker exec "$docker_id" sh -c 'chmod +x /opt/simulation_vba/simulation_vba/vba_emu.py 2>/dev/null || true'

echo "[*] Selecting Python runtime with SimulationVBA dependencies..."
vm_python="$(docker exec "$docker_id" sh -c '
    cd /opt/simulation_vba || exit 1
    for py in python3 python pypy3 pypy python2; do
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
extra_opts=()

if [[ $# -ge 3 && "${2:-}" == "-i" ]]; then
    entry="-i $3"
    extra_opts=("${@:4}")
elif [[ $# -eq 2 ]]; then
    json="-o $container_json"
    json_file="$2"
    extra_opts=()
elif [[ $# -ge 2 ]]; then
    extra_opts=("${@:2}")
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
extra_opts_str=""
if [ ${#extra_opts[@]} -gt 0 ]; then
    for opt in "${extra_opts[@]}"; do
        extra_opts_str+=" $(printf "%q" "$opt")"
    done
fi
docker exec "$docker_id" sh -c "cd /opt/simulation_vba && PYTHONPATH=/opt/simulation_vba:/opt/simulation_vba/simulation_vba/core $vm_python simulation_vba/vba_emu.py -s --ioc --jit $json $entry$extra_opts_str '$container_input'"

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
