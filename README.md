# Simulation VBA

Simulation VBA is a VBA macro emulation and analysis tool for Microsoft Office documents. It is designed to help inspect suspicious VBA macros, deobfuscate strings, simulate macro execution, collect IOCs, and dump files created by the macro during emulation.

The tool supports Office Open XML files such as `.xlsm`, `.docm`, `.pptm`, and related formats. It also includes Excel workbook context handling for common macro hiding techniques, such as payload fragments stored in shapes, UserForm properties, named ranges, selected cells, and worksheet strings.

## Main Features

- Parse and emulate VBA macros from Office files.
- Detect auto-execution entry points such as `Auto_Open`.
- Track suspicious actions such as file writes, shell execution, environment access, and dropped files.
- Resolve Excel OOXML runtime context, including shapes, defined names, cells, selections, and UserForm strings.
- Decode and reconstruct payloads that are split across multiple workbook locations.
- Save dropped files and artifact archives to the directory where the command is run.
- Run analysis inside a fresh Docker container with networking disabled.

## Requirements

Recommended usage requires Docker.

The Docker wrapper creates a new disposable container for each analysis run. The local source tree is copied into the container each time, so local code changes are used immediately.

## Basic Usage

From the repository root:

```bash
chmod +x docker/simulation_vba.sh
./docker/simulation_vba.sh path/to/sample.xlsm
```

Example:

```bash
./docker/simulation_vba.sh ~/samples/invoice-42369643.xlsm
```

## Output

The tool prints macro parsing results, emulation logs, recorded actions, and possible IOCs to the terminal.

If the macro drops files during emulation, artifacts are copied to the directory where the command was run:

```text
sample.xlsm_artifacts/
sample.xlsm_artifacts.zip
```

The artifact zip is password-protected with:

```text
infected
```

To inspect the archive:

```bash
unzip -l sample.xlsm_artifacts.zip
```

To extract it:

```bash
unzip sample.xlsm_artifacts.zip
```

When prompted for the password, enter:

```text
infected
```

## JSON Report

A JSON report can be written by passing an output filename as the second argument:

```bash
./docker/simulation_vba.sh path/to/sample.xlsm report.json
```

The JSON file is written to the directory where the command was run, unless an absolute path is provided.

## Custom Entry Point

To start emulation from a specific macro entry point:

```bash
./docker/simulation_vba.sh path/to/sample.xlsm -i Auto_Open
```

With JSON output:

```bash
./docker/simulation_vba.sh path/to/sample.xlsm report.json -i Auto_Open
```

## Docker Options

The wrapper supports a few environment variables:

```bash
SIMULATIONVBA_DOCKER_PULL=0 ./docker/simulation_vba.sh sample.xlsm
```

Do not pull the base Docker image before running.

```bash
SIMULATIONVBA_DOCKER_KEEP=1 ./docker/simulation_vba.sh sample.xlsm
```

Keep the container after analysis for debugging.

```bash
SIMULATIONVBA_DOCKER_IMAGE=image_name ./docker/simulation_vba.sh sample.xlsm
```

Use a custom Docker image as the runtime base.

## Optional Local Installation

The tool can also be installed as a Python package:

```bash
pip install -e .
```

After installation, the command-line entry point is:

```bash
simulation_vba path/to/sample.xlsm
```

For common macro analysis options:

```bash
simulation_vba -s --iocs --jit path/to/sample.xlsm
```

Docker usage is recommended for suspicious files.

## Safety Notes

Analyze untrusted Office files only in an isolated environment.

The Docker wrapper starts the container with networking disabled. This helps reduce risk during macro emulation, but it is not a replacement for a dedicated malware analysis environment.

Do not open suspicious Office documents directly in Microsoft Office on your main system.

## Typical Workflow

1. Run the sample through the Docker wrapper.
2. Review recorded actions in the terminal output.
3. Check for suspicious behavior such as file writes, shell commands, and decoded payloads.
4. Inspect any dumped artifacts in the generated artifact directory or zip file.
5. Use the JSON report if structured output is needed.

Example:

```bash
./docker/simulation_vba.sh sample.xlsm report.json
unzip -l sample.xlsm_artifacts.zip
```

## Project Layout

```text
docker/simulation_vba.sh        Docker wrapper
simulation_vba/vba_emu.py       Main emulator entry point
simulation_vba/core/            VBA parser, emulator, and runtime logic
simulation_vba/core/ooxml_context.py  Excel OOXML context resolver
setup.py                        Optional Python package installation
requirements.txt                Python dependencies
```

## Notes

Simulation VBA is intended for static and emulated analysis of VBA macro behavior. It does not execute macros inside Microsoft Office. Results should be reviewed manually, especially for heavily obfuscated or unsupported VBA features.
