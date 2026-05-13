# Usage

## Install

=== "PyPI"

    ```bash
    pip install vbalidator
    ```

=== "Docker (GHCR)"

    ```bash
    docker pull ghcr.io/twobeass/vbalidator:latest
    docker run --rm \
      -v "$PWD/MyModules:/workspace" \
      ghcr.io/twobeass/vbalidator:latest \
      /workspace --host excel
    ```

=== "From source"

    ```bash
    git clone https://github.com/twobeass/VBAlidator
    cd VBAlidator
    pip install -e ".[dev]"
    pytest                       # 259 tests
    ```

## CLI

```bash
vbalidator <input> [options]
```

`<input>` is either a single `.bas` / `.cls` / `.frm` file or a folder
that is walked recursively.

### Options

| Flag | Default | Purpose |
|------|---------|---------|
| `--host {excel,word,access,outlook,visio,mscomctl,msforms,scripting,vbscript_regexp,wscript_shell,shell_application}` | _none_ | Auto-load the bundled host model. The five Office hosts (excel/word/access/visio/outlook) need to be set explicitly; the six COM-companion stubs (mscomctl/msforms/scripting/vbscript_regexp/wscript_shell/shell_application) **auto-layer** when the scan set mentions their ProgID / namespace — explicit `--host` rarely needed for those. See [Configuration → Bundled host models](Configuration.md#bundled-host-models). |
| `--model PATH` | `vba_model.json` if present | Custom JSON object model. Layered on top of the std model and any `--host` model. |
| `--define KEY=VAL,KEY2=VAL2` | _none_ | Conditional-compilation constants. Override `WIN64` / `VBA7` to force 32-bit mode. |
| `--score-threshold N` | `90` | Minimum score for a clean exit. |
| `--strict` / `--no-strict` | `--strict` | Whether `severity=warning` findings count toward the gating score. Errors always do. |
| `--roundtrip` | off | Cross-check via the actual VBE compiler. Windows + Office + pywin32 only; degrades gracefully off-platform. |
| `--quiet` | off | Suppress per-issue output, print summary only. |
| `--output PATH` | `vba_report.json` | Where to write the JSON v2 report. |

### Exit codes

| Code | Meaning |
|------|---------|
| 0 | `compile_safe == True` and `score ≥ threshold` |
| 1 | Score below threshold, or at least one error |
| 2 | Input path does not exist |
| 3 | Pipeline crash |
| 4 | Could not write the JSON report |

### Examples

```bash
# Smoke a folder of Excel modules
vbalidator ./MyModules --host excel

# Force 32-bit Office assumptions
vbalidator ./MyModules --define WIN64=False,VBA7=False

# CI gate — fail the build below 95
vbalidator ./vba --host excel --score-threshold 95

# Pure-error gate (no warning noise)
vbalidator ./vba --host excel --no-strict --quiet

# Static + dynamic cross-check (Windows + Office only)
vbalidator ./vba --host excel --roundtrip
```

## Python API

```python
from vbalidator import precheck, PrecheckResult

result: PrecheckResult = precheck(
    source="./MyModules",        # str | Path | inline source
    host="excel",                # excel|word|access|outlook|visio|mscomctl|msforms|scripting|vbscript_regexp|wscript_shell|shell_application|None
    model_path="my.json",        # extra custom model
    defines={"WIN64": False},
    strict=True,                 # warnings count toward score
    module_type=None,            # override for inline strings
    roundtrip=False,             # Windows + Office only
)

result.score          # 0..100
result.compile_safe   # True / False
result.errors         # list[Issue]
result.warnings       # list[Issue]
result.info           # list[Issue]
result.issues         # full list, normalised
result.json()         # canonical JSON v2 report

bool(result)          # truthy when compile_safe
```

For inline strings without an associated path, prefer
`precheck_source(code, name="<my-snippet>", host="excel")`.

### Issue shape (JSON v2)

```json
{
  "rule_id": "VBA001",
  "severity": "error",
  "category": "name_resolution",
  "file": "Module1.bas",
  "line": 42,
  "column": 0,
  "message": "Undefined identifier 'tpyo' in 'DoStuff'."
}
```

`severity` is one of `error`, `warning`, `info`, or `compile_verified`
(round-trip).

### Rule IDs

Every emitted finding carries a stable `rule_id` documented at
[Rules](rules/index.md). Use them in CI ignore lists:

```bash
vbalidator ./vba --host excel | jq '.issues[] | select(.rule_id != "VBA320")'
```

## Configuration files

`vba_model.json` next to the working directory is auto-loaded if no
`--model` is given. See [Configuration](Configuration.md) for the
schema and how to generate one with the bundled
`tools/VBA_Model_Exporter.bas`.
