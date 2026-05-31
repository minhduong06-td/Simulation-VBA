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
- Extract `AddFromString` dynamically injected VBA code to `addfromstring_N.vba`.
- Extract shellcode from `Array(...)` patterns to `shellcode.bin` (with API call detection: `CreateProcessA`, `VirtualAllocEx`, `WriteProcessMemory`, `CreateRemoteThread`).
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

### CLI wrapper (recommended)

The `dockervba` CLI wrapper at `~/tools/bin/dockervba` provides argument validation and a clean interface. Make sure it is on your `PATH`:

```bash
export PATH="$HOME/tools/bin:$PATH"
```

Or alias the script directly:

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

All examples assume `dockervba` is on your `PATH`. Replace with `./docker/simulation_vba.sh` if not.

### Simulate an Office file

```bash
dockervba sample.xlsm
```

### Deobfuscate and simulate an HTA or VBScript

```bash
dockervba --deob simulate sample.hta
```

Options can also appear after the file:

```bash
dockervba sample.hta --deob simulate
```

### Deobfuscate and simulate an Office file

```bash
dockervba --deob simulate sample.xlsm
```

**Note:** Only `--deob simulate` is supported. `--deob emulate` is not supported and will be rejected with an error.

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

Files dropped during emulation are copied to `<input>_artifacts/` and `<input>_artifacts.zip`:

```text
<input>_artifacts/
<input>_artifacts.zip
```

### Artifact types

| File | Description |
|------|-------------|
| `*.hta`, `*.vbs`, `*.exe`, etc. | Files written by the macro via `Open` / `Put` / `WriteText` |
| `addfromstring_N.vba` | VBA code dynamically injected via `xlmodule.CodeModule.AddFromString` |
| `shellcode.bin` | Shellcode bytes extracted from `Array(...)` patterns in deobfuscated code |

### Example: XLSM analysis

```
invoice-42369643.xlsm_artifacts/
└── LwTHLrGh.hta            (10631 bytes — decoded HTA payload)
```

### Example: HTA `--deob simulate` analysis

```
stage2.hta_artifacts/
├── addfromstring_1.vba      (636 bytes — injected VBA code)
└── shellcode.bin            (416 bytes — shellcode from Array)
```

### Password-protected zip

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
- Extracted artifacts (`addfromstring_N.vba`, `shellcode.bin`) are written to disk for manual inspection only — never executed automatically.
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

# Deobfuscate HTA (extracts AddFromString + shellcode)
dockervba --deob simulate path/to/stage2.hta

# File-first ordering also works
dockervba path/to/stage2.hta --deob simulate

# Verify artifacts
find . -maxdepth 3 -name "_artifacts" -type d

# Clean repo artifacts only
dockervba --clean-artifacts

# --deob emulate is rejected
dockervba --deob emulate path/to/sample.hta
# → Unsupported --deob mode: emulate. Supported: simulate
```

## Project Layout

```text
docker/simulation_vba.sh        Docker wrapper (also callable directly)
docker/Dockerfile               Docker image definition
setup.py                        Optional pip install
requirements.txt                Python dependencies
simulation_vba/vba_emu.py       Main emulator entry point
simulation_vba/core/            VBA parser, emulator, and runtime logic
  vba_library.py                AddFromString hook, WriteProcessMemory, VBA functions
  vba_context.py                Shellcode accumulator, out_dir, file I/O context
  deobfuscation.py              HTA extraction, deobfuscation, stubbed simulation engine
  ooxml_context.py              Excel OOXML context resolver
```

Simulation VBA is intended for static and emulated analysis of VBA macro behavior. Results should be reviewed manually, especially for heavily obfuscated or unsupported VBA features.
