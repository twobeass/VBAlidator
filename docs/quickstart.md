# Quickstart

Get VBAlidator running against your VBA codebase in under five minutes.

## Install

=== "PyPI"

    ```bash
    pip install vbalidator
    ```

=== "Docker"

    ```bash
    docker pull ghcr.io/twobeass/vbalidator:latest
    ```

=== "From source"

    ```bash
    git clone https://github.com/twobeass/VBAlidator
    cd VBAlidator
    pip install -e ".[dev]"
    ```

## First scan

Export your VBA modules from the host (`File → Export Module…` in the
VBE) into a folder. Then run:

```bash
vbalidator ./MyModules --host excel
```

Sample output:

```text
VBAlidator: scanning ./MyModules (host=excel)
MyModules/Module1.bas:42: ERROR  [VBA001]  Undefined identifier 'tpyo' in 'DoStuff'.
MyModules/Module1.bas:1:  WARNING [VBA320]  Module 'Module1' is missing `Option Explicit`.

Files scanned : 12
Errors        : 1
Warnings      : 1
Info          : 0
Confidence    : 77 / 100  (needs fixes)
Report saved  : vba_report.json
```

The CLI exits **non-zero** when the score is below the threshold
(default 90) or any error is present, so you can drop it directly
into any CI step.

## CI usage

```yaml
# GitHub Actions
- run: pip install vbalidator
- run: vbalidator ./vba --host excel --score-threshold 90
```

```yaml
# GitLab CI
vbalidator:
  image: ghcr.io/twobeass/vbalidator:latest
  script:
    - vbalidator . --host excel
```

## Python API

```python
from vbalidator import precheck

result = precheck(
    source="./MyModules",        # file, dir, or inline source string
    host="excel",                # excel | word | access | outlook
    defines={"WIN64": True},     # conditional-compilation constants
    strict=True,                 # warnings count toward gating
    roundtrip=False,             # set True on Windows for VBE compile check
)

print(result.score, result.compile_safe)
print(result.json())             # canonical JSON v2 report
```

See [AI Pipeline Integration](ai-integration.md) for end-to-end
recipes with LangChain, the Anthropic SDK and OpenAI Agents.

## Common flags

| Flag | Purpose |
|------|---------|
| `--host excel` | Auto-load the bundled Excel object model |
| `--model my.json` | Use a custom JSON object model |
| `--define WIN64=False` | Override conditional-compilation constants |
| `--score-threshold 80` | Lower the CI gate from 90 to 80 |
| `--no-strict` | Warnings stop counting against the score |
| `--roundtrip` | Cross-check via the actual VBE (Windows + Office only) |
| `--quiet` | Suppress per-issue output, print summary only |
| `--output rep.json` | Write the JSON v2 report to a path |

Full reference: [Usage Guide](Usage.md).
