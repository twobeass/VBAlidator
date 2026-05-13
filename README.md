# VBAlidator

> **Premium VBA static analyser & compile-safety prechecker for AI-generated VBA.**
> Drop-in behind any LLM-VBA generator — get a deterministic 0–100
> confidence score and a stable JSON report before the code ever
> reaches a workbook.

[![PyPI version](https://img.shields.io/pypi/v/vbalidator)](https://pypi.org/project/vbalidator/)
[![CI](https://github.com/twobeass/VBAlidator/actions/workflows/ci.yml/badge.svg)](https://github.com/twobeass/VBAlidator/actions/workflows/ci.yml)
[![Docker](https://img.shields.io/badge/ghcr.io-twobeass%2Fvbalidator-blue?logo=docker)](https://github.com/twobeass/VBAlidator/pkgs/container/vbalidator)
[![Docs](https://img.shields.io/badge/docs-twobeass.github.io-blue)](https://twobeass.github.io/VBAlidator/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

VBAlidator parses `.bas` / `.cls` / `.frm` files through a real
lexer / preprocessor / parser / analyzer pipeline, applies **41
documented rules** (see [the catalogue](docs/rules/index.md)), and
returns a verdict CI pipelines and AI agents can act on directly.

## What it catches

The full list lives at [`docs/rules/index.md`](docs/rules/index.md).
The high-impact subset that AI generators trip over most often:

- **VBA001** undefined identifiers, **VBA002** missing object members
- **VBA006** wrong argument count, **VBA007** ByRef type mismatches
- **VBA210** `Set` on scalar, **VBA211** missing `Set` on object assignment
- **VBA221–VBA224** Property Get/Let/Set arity & semantics
- **VBA230 / VBA231** non-constant Const initialisers
- **VBA240** arithmetic between string and numeric literal
- **VBA300** missing `PtrSafe` on 64-bit Office Declares
- **VBA320** missing `Option Explicit` (warning)
- **VBA330** incomplete `Implements` of an interface
- **VBA340 / VBA341** `RaiseEvent` without matching `Event` declaration / wrong arity
- **VBA_LEX001 / VBA_LEX002** unrecognised characters & malformed date literals
- **VBA_RT001** *(optional)* errors caught by the actual VBE compiler via Office round-trip

## Install

```bash
pip install vbalidator
```

…or grab the multi-arch image:

```bash
docker pull ghcr.io/twobeass/vbalidator:latest
```

## In one CLI call

```bash
vbalidator ./MyModules --host excel
```

```text
MyModules/Module1.bas:42: ERROR  [VBA001]  Undefined identifier 'tpyo' in 'DoStuff'.
MyModules/Module1.bas:1:  WARNING [VBA320]  Module 'Module1' is missing `Option Explicit`.

Files scanned : 12
Errors        : 1
Warnings      : 1
Confidence    : 77 / 100  (needs fixes)
Report saved  : vba_report.json
```

Exit code is `1` when the score is below the gate (`--score-threshold`,
default 90) or there is at least one error.

## In one Python call

```python
from vbalidator import precheck

result = precheck("Module1.bas", host="excel")

if result.compile_safe:
    deploy(result)
else:
    for err in result.errors:
        print(f"{err['rule_id']}: {err['message']}")

print(f"score = {result.score} / 100")
print(result.json())   # canonical JSON v2 report
```

`PrecheckResult` exposes `errors`, `warnings`, `info`, `issues`, and
the canonical `.json()` report. It's also truthy when `compile_safe`,
so `if precheck(...): ...` is idiomatic.

See [docs/ai-integration.md](docs/ai-integration.md) for full
recipes (Anthropic SDK, OpenAI Agents, LangChain, GitHub Actions).

## Optional dynamic verification

When a Windows host with Office is available, cross-check the static
verdict against the actual VBE compiler:

```bash
vbalidator ./MyModules --host excel --roundtrip
```

```python
result = precheck("Module1.bas", host="excel", roundtrip=True)
```

Compile errors VBE itself reports come back with
`severity='compile_verified'` and rule_id `VBA_RT001`. On non-Windows
hosts the call degrades gracefully to a single info-level notice
instead of crashing.

## Bundled host models & auto-layering

`--host excel|word|access|outlook|visio` auto-loads the matching Office
host model from `src/models/`. Excel/Word/Access/Visio are
**full-fidelity** exports of the real Office type libraries
(1–3 MB each, ~1000 classes apiece, ~5000 globals/classes in
Excel alone); Outlook is a hand-curated stub (the Trust-Center
AccessVBOM path is GPO-blocked on managed installs).

In addition, six companion stubs **auto-layer** without an explicit
`--host` flag whenever any scanned file mentions their ProgID /
namespace:

| Stub | Triggers on | Top classes |
|---|---|---|
| `mscomctl` | `.frm` referencing `ComctlLib` | TreeView, ListView, Toolbar, ProgressBar |
| `msforms` | source mentioning `MSForms.` | UserForm, CommandButton, TextBox, Frame |
| `scripting` | `Scripting.Dictionary` / `Scripting.FileSystemObject` | Dictionary, FSO, Drive, Folder, File, TextStream |
| `vbscript_regexp` | `VBScript.RegExp` | RegExp, Match, MatchCollection, SubMatches |
| `wscript_shell` | `WScript.Shell` | Shell, WshExec, WshEnvironment |
| `shell_application` | `Shell.Application` | Shell.Application, Shell.Folder, Shell.FolderItem |

So `--host excel` covers the vast majority of real-world spreadsheet
VBA — TreeView/UserForm/Dictionary/RegExp code resolves
automatically. Run `vbalidator --help` for the full `--host` choice
list.

## Documentation

The full site is at **<https://twobeass.github.io/VBAlidator/>** —
deployed automatically from `main` by `.github/workflows/docs.yml`.

In-repo:

- [Quickstart](docs/quickstart.md) — install → scan → ship
- [GitHub Actions integration](docs/github-actions.md) — copy-paste
  workflow recipes (path-filter, JSON artifact, PR annotations,
  multi-host matrix, step summary, status badge)
- [AI pipeline integration](docs/ai-integration.md) — patterns for
  Claude, OpenAI Agents, LangChain
- [Usage](docs/Usage.md) — full CLI / Python API reference
- [Configuration](docs/Configuration.md) — host models, custom models,
  conditional-compilation defaults, heuristics
- [Architecture](docs/Architecture.md) — pipeline internals
- [Rule catalogue](docs/rules/index.md) — all 41 rules with examples
- [CI / CD](docs/ci-cd.md) — workflows, release lifecycle, branch
  protection
- [UAT walkthrough](docs/uat.md) — section-by-section human validation
  script for the current branch
- [TODO](TODO.md) — open items that need a Windows machine, live
  PyPI/GHCR/Pages access, or future code work
- [Roadmap](docs/roadmap.md) — what's done, what's queued

## Development

```bash
git clone https://github.com/twobeass/VBAlidator
cd VBAlidator
pip install -e ".[dev]"
pytest                              # 259 tests
ruff check src tests                # lint
python tools/generate_rule_docs.py  # refresh docs/rules/ after a rule change
```

PR titles must follow [Conventional Commits] (`feat:` / `fix:` / `docs:` …).
The release pipeline derives the next semver bump from your commit
history.

[Conventional Commits]: https://www.conventionalcommits.org/

## License

MIT.
