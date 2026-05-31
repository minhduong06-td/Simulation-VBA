# Simulation VBA

Simulation VBA is a VBA macro emulation and analysis tool for Microsoft Office documents. It parses and emulates VBA macros, deobfuscates payloads, records IOCs and actions, and extracts dropped artifacts — all inside an isolated Docker container with networking disabled.

Supports Office Open XML formats (`.xlsm`, `.docm`, `.pptm`) and standalone HTA/VBScript files.

## Features

- Parse and emulate VBA macros from Office files.
- Detect auto-execution entry points (`Auto_Open`, etc.).
- Track suspicious actions: file writes, `Shell`, registry access, process injection, `CreateObject`.
- Resolve Excel OOXML runtime context: shapes, defined names, cells, selections, UserForm strings.
- Decode payloads split across multiple workbook locations.
- Deobfuscate and simulate HTA / VBScript with dangerous actions stubbed.
- Save dropped files and artifact archives to the working directory.
- Run inside a fresh Docker container with `--network none`.

## Requirements

- Docker
- Bash (Linux or WSL)

Python 3 is only needed for local development (optional). The Docker wrapper uses Python 3 inside the container automatically.

## Installation

```bash
git clone <repo-url> Simulation-VBA
cd Simulation-VBA
chmod +x docker/simulation_vba.sh
```

Optional alias:

```bash
alias dockervba="$PWD/docker/simulation_vba.sh"
```

## Docker Image Behavior

- **First run**: if the local image `md06/simulation:latest` does not exist, the script builds it from `docker/Dockerfile`.
- **Subsequent runs**: the cached image is reused — no rebuild.
- **Each run**: a temporary container is created, the current source tree is copied in, and the container is removed when done.
- **Source changes**: no need to rebuild. Source is copied fresh each run, so code edits take effect immediately.
- **Network**: containers are started with `--network none` for safety.

Use `--rebuild` to force a fresh image build (e.g., after changing `docker/Dockerfile` or `requirements.txt`).

## Usage

All examples assume `dockervba` is aliased. Replace with `./docker/simulation_vba.sh` if not.

### Simulate an Office file

```bash
dockervba sample.xlsm
```

### Deobfuscate and simulate (Office)

```bash
dockervba --deob simulate sample.xlsm
```

### Deobfuscate and simulate an HTA or VBScript

```bash
dockervba --deob simulate sample.hta
```

Options can also appear after the file:

```bash
dockervba sample.hta --deob simulate
```

### Rebuild the Docker image

```bash
dockervba --rebuild
```

Or rebuild and run in one step:

```bash
dockervba --rebuild sample.xlsm
```

### Clean artifacts only (no Docker)

```bash
dockervba --clean-artifacts
```

Removes `*_artifacts/`, `*_artifacts.zip`, `__pycache__`, `*.pyc` from the repo.

### Full cleanup (Docker + artifacts)

```bash
dockervba --clean
```

Removes:
- Containers `simulation-vba-run-*`
- Temp images `simulation-vba-temp:*`
- Cached image `md06/simulation:latest`
- Repo build artifacts

### Custom entry point

```bash
dockervba sample.xlsm -i Auto_Open
```

### JSON report

```bash
dockervba sample.xlsm report.json
```

With custom entry point:

```bash
dockervba sample.xlsm report.json -i Auto_Open
```

## Output / Artifacts

Emulation logs, recorded actions, and IOCs are printed to the terminal.

If the macro drops files during emulation, they are copied to:

```text
<input>_artifacts/
<input>_artifacts.zip
```

Example:

```
invoice-42369643.xlsm_artifacts/LwTHLrGh.hta
```

The zip is password-protected with `infected`:

```bash
unzip -l sample.xlsm_artifacts.zip
unzip sample.xlsm_artifacts.zip   # password: infected
```

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `SIMULATIONVBA_DOCKER_IMAGE` | `md06/simulation:latest` | Docker image tag |
| `SIMULATIONVBA_DOCKER_PULL` | `0` | Pull image before each run |
| `SIMULATIONVBA_DOCKER_KEEP` | `0` | Keep container after run for debugging |

## Safety Notes

- Analyze untrusted files only in an isolated environment (VM, sandbox).
- The Docker wrapper starts containers with `--network none`.
- Dangerous actions (`Shell`, `CreateObject`, registry, process injection) are logged and stubbed — never executed for real.
- The `--deob simulate` path is non-destructive; unsupported statements are preserved in output.
- Do not open suspicious Office documents in Microsoft Office on your main system.

## Developer Notes

- Source is copied into the container every run. Edit code, re-run — no rebuild needed.
- Rebuild image only when `docker/Dockerfile` or `requirements.txt` changes:

  ```bash
  dockervba --rebuild
  ```

- To check for leftover Docker resources:

  ```bash
  docker ps -a | grep -iE 'simulation|vba'
  docker images | grep -iE 'simulation-vba-temp|md06/simulation'
  ```

## Testing Checklist

```bash
# Full cleanup
dockervba --clean

# Simulate Office file
dockervba path/to/sample.xlsm

# Deobfuscate HTA
dockervba --deob simulate path/to/sample.hta

# Clean repo artifacts only
dockervba --clean-artifacts
```

## Project Layout

```text
docker/simulation_vba.sh        Docker wrapper
docker/Dockerfile               Docker image definition
setup.py                        Optional pip install
requirements.txt                Python dependencies
simulation_vba/vba_emu.py       Main emulator entry point
simulation_vba/core/            VBA parser, emulator, and runtime logic
simulation_vba/core/ooxml_context.py  Excel OOXML context resolver
```

Simulation VBA is intended for static and emulated analysis of VBA macro behavior. Results should be reviewed manually, especially for heavily obfuscated or unsupported VBA features.
