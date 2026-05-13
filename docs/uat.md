# Human UAT — branch validation walkthrough

This walkthrough exists so a human reviewer can sign off on every
single capability the `claude/improve-vba-precompiler-3YcZp` branch
ships. It mirrors the Phase 0–5 roadmap section by section and ends
with the items that genuinely require a Windows + Office machine.

Every step lists:

- **What** to run.
- **Expected outcome** with the exact text / numbers to look for.
- **Tick when verified** so you can keep score as you go.

If anything diverges from the expected output, jot it down — the
analyser is deterministic, so divergence almost always points to an
actual regression.

---

## Section 0 — Setup

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `git fetch origin && git switch claude/improve-vba-precompiler-3YcZp` | Clean checkout. |
| ☐ | `python --version` | `>=3.10` (CI tests against 3.10–3.13; Py3.9 was dropped after upstream EOL Oct 2025). |
| ☐ | `python -m pip install -e ".[dev]"` | Installs `vbalidator` + `pytest` + `ruff`. |
| ☐ | `vbalidator --version` *(if shipped)* or `python -c "import src; print(src.__version__)"` | Prints the current version (`0.1.1` until the first semantic-release tag). |

> If you're on macOS / Linux without a venv, prefer `pip install -e ".[dev]" --user` or `pipx install -e .` to keep system Python clean.

---

## Section 1 — Test suite

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `pytest -ra` | **152 passed** in ~2 s. No skips on Linux/macOS; the three `roundtrip_off_platform` tests skip only on Windows. |
| ☐ | `pytest --cov=src` then `coverage report` | Total coverage **≥ 83 %**, with `src/scoring.py` and `src/rules.py` at 100 %. |
| ☐ | `ruff check src tests tools` | `All checks passed!` |
| ☐ | `python tools/generate_rule_docs.py` | `docs/rules/: 41 rules, 0 files updated.` (idempotent) |

If the rule generator reports `> 0 files updated`, **fail UAT** —
something diverged between `src/rules.py` and the committed catalogue.

---

## Section 2 — CLI baseline

The fixtures under `tests/demo/` deliberately contain compile errors
and `tests/samples/valid_code/` is clean. They are the canonical UAT
inputs.

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `vbalidator tests/samples/valid_code/valid_sample.bas --quiet --no-strict --output /tmp/clean.json` then `echo $?` | Exit code **0**. JSON shows `"score": 100, "compile_safe": true`. |
| ☐ | `vbalidator tests/demo --quiet --output /tmp/dirty.json` then `echo $?` | Exit code **1**. JSON shows ≥ 8 errors, `compile_safe: false`, score 0. |
| ☐ | `python -c "import json; d=json.load(open('/tmp/dirty.json')); print(d['version'], d['summary'])"` | `2.0` plus a summary dict with `errors`, `warnings`, `info`, `score`, `compile_safe`, `files_scanned`, `issues_total`. |
| ☐ | `python -c "import json; d=json.load(open('/tmp/dirty.json')); print({i['rule_id'] for i in d['issues']})"` | At least `{'VBA001', 'VBA002', 'VBA003', 'VBA005', 'VBA320'}` are present. |

---

## Section 3 — `--host` bundled models

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `vbalidator tests/samples/valid_code/valid_sample.bas --host excel --quiet --no-strict` | Exit 0, score 100. |
| ☐ | Create `/tmp/excel-test.bas` containing<br>`Attribute VB_Name = "M"`<br>`Option Explicit`<br>`Sub S(): Dim wb As Workbook: Set wb = ActiveWorkbook: wb.Save: End Sub` and run `vbalidator /tmp/excel-test.bas --host excel --quiet` | Exit 0. **No** `VBA001` (`Workbook`/`ActiveWorkbook`/`.Save` resolve via `models/excel.json`). |
| ☐ | Same file with `--host word` | At least one `VBA001`/`VBA002` (Word doesn't know `Workbook`). |
| ☐ | Repeat with the Word, Access, Outlook examples in `docs/quickstart.md` | Each host resolves its own types; cross-host invocations fail loudly. |

---

## Section 4 — `--score-threshold` and `--strict / --no-strict`

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | Drop `Option Explicit` from a clean module and run with default flags | Exit code 0 still (warning costs 3 pts → score 97 ≥ 90), but `VBA320` shows in the report. |
| ☐ | Same input with `--score-threshold 99` | Exit code **1** (97 < 99 fails the gate). |
| ☐ | Same input with `--no-strict` | Score back at 100, exit 0 (warning excluded from gating). |
| ☐ | Module with one undefined identifier, `--no-strict` | Exit 1 still — errors always block, regardless of strictness. |

---

## Section 5 — Python API

Inside a Python REPL or a scratch file:

```python
from vbalidator import precheck, precheck_source, PrecheckResult

# Inline source
r = precheck_source("""
Attribute VB_Name = "M"
Option Explicit
Sub S(): End Sub
""")
assert r.compile_safe and r.score == 100, (r.score, r.compile_safe)
assert isinstance(r, PrecheckResult)

# Truthy iff compile_safe
assert bool(r)

# Errors disjoint from warnings
assert {e["message"] for e in r.errors} & {w["message"] for w in r.warnings} == set()

# JSON v2
j = r.json()
assert j["version"] == "2.0"
assert j["summary"]["score"] == 100

# Path input
r2 = precheck("tests/demo")
assert not r2.compile_safe, "demo must be dirty"
```

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | All asserts above pass without error |  |
| ☐ | `r.errors`, `r.warnings`, `r.info` are disjoint lists |  |
| ☐ | `r.json()["version"] == "2.0"` |  |
| ☐ | `bool(precheck("tests/demo")) is False` |  |
| ☐ | `bool(precheck("tests/samples/valid_code/valid_sample.bas")) is True` |  |

---

## Section 6 — `vba_model.json` auto-load

This was lost in the Phase 4 rewrite and restored in this branch.

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `mkdir /tmp/uat-auto && cd /tmp/uat-auto` |  |
| ☐ | `printf 'Attribute VB_Name = "M"\nOption Explicit\nSub S(): Dim x As Long: x = MyCustomGlobal: End Sub\n' > /tmp/uat-auto/M.bas` | Module references `MyCustomGlobal` (not in any standard model). |
| ☐ | `vbalidator /tmp/uat-auto --quiet` | Exit 1 — `VBA001` for `MyCustomGlobal`. |
| ☐ | `printf '{"globals":{"MyCustomGlobal":{"type":"Long"}}}' > /tmp/uat-auto/vba_model.json` | Drop a custom model next to the input. |
| ☐ | `vbalidator /tmp/uat-auto --quiet` (re-run) | Exit 0 — `vba_model.json` is auto-loaded; `MyCustomGlobal` now resolves. |
| ☐ | Move the model up one folder and run from there | Same — the search reaches `<input_dir>` → `<input_file_dir>` → `<cwd>`. |
| ☐ | Pass `--model /tmp/explicit.json` (different content) | The explicit model wins; the auto-detected one is ignored. |

---

## Section 7 — Round-trip verification (Linux fallback)

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `vbalidator tests/samples/valid_code/valid_sample.bas --quiet --no-strict --roundtrip` (on Linux) | Exit 0. JSON contains a single `VBA_RT000` info issue with `severity: "info"` saying "Round-trip verification unavailable". `compile_safe` stays True. |
| ☐ | `python -c "from src.roundtrip import is_available; print(is_available())"` | `False` on Linux, `True` on Windows with pywin32 installed. |

The full Windows path is in **Section 12**.

---

## Section 8 — Each shipped rule with a fixture

The compile-error fixtures double as UAT inputs. Run each manually:

```bash
for cat in tests/samples/compile_errors/*/; do
  echo "=== $cat ==="
  vbalidator "$cat" --quiet --output /tmp/cat.json
  python -c "
import json
d = json.load(open('/tmp/cat.json'))
ids = sorted({i['rule_id'] for i in d['issues']})
print('   exit-rule_ids:', ids)
"
done
```

| ☐ | Category folder | Expected rule_id present |
|---|-----------------|--------------------------|
| ☐ | `argument_mismatch/` | `VBA006` |
| ☐ | `byref_mismatch/` | `VBA007` |
| ☐ | `const_expression/` | `VBA230` and/or `VBA231` |
| ☐ | `date_literal/` | `VBA_LEX002` |
| ☐ | `declare_ptrsafe/` | `VBA300` |
| ☐ | `duplicate_declaration/` | `VBA003` |
| ☐ | `enum_uniqueness/` | `VBA310` |
| ☐ | `erase_target/` | `VBA105` or `VBA106` |
| ☐ | `fixed_length_string/` | `VBA250` |
| ☐ | `jump_target/` | `VBA201` |
| ☐ | `member_access/` | `VBA002` and/or `VBA004` |
| ☐ | `operator_type/` | `VBA240` |
| ☐ | `property_arity/` | `VBA221`, `VBA222`, `VBA223` or `VBA224` |
| ☐ | `raise_event/` | `VBA340` |
| ☐ | `redim_target/` | `VBA101`, `VBA102` or `VBA103` |
| ☐ | `set_vs_let/` | `VBA210` and/or `VBA211` |
| ☐ | `syntax_errors/` | `VBA010` |
| ☐ | `type_mismatch/` | `VBA005` and/or related |
| ☐ | `undefined_identifier/` | `VBA001` |
| ☐ | `unreachable_code/` | `VBA009` |

If any category surfaces a different / extra rule, that's information —
record it. The catalogue in `docs/rules/index.md` documents what each
ID means.

---

## Section 9 — Documentation site (MkDocs)

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `pip install ".[docs]"` | Installs MkDocs + Material. |
| ☐ | `mkdocs build --strict` | Builds without warnings. Output in `./site/`. |
| ☐ | `mkdocs serve` and open <http://localhost:8000> | Material theme, all top-level nav entries (Home, Quickstart, AI Pipeline Integration, User Guide, Rule catalogue, CI/CD, Roadmap, Changelog) load. |
| ☐ | Visit `Rule catalogue → Overview` | 35 rule pages listed in the table. |
| ☐ | Click a few rule pages (e.g. `VBA210`, `VBA300`) | Each shows Description / Failing example / Compliant example / How to fix. |
| ☐ | Open *Quickstart* and *AI Pipeline Integration* | Code blocks render with copy buttons (Material `content.code.copy`). |

---

## Section 10 — Docker image

Requires Docker locally.

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `docker build -t vbalidator:uat .` | Builds two stages (`builder`, `runtime`); final image ≤ 200 MB. |
| ☐ | `docker run --rm vbalidator:uat --help` | Prints the same help as the native CLI. |
| ☐ | `docker run --rm -v "$PWD/tests/demo:/workspace" vbalidator:uat /workspace --quiet` then `echo $?` | Exit 1, score 0 (same as native run). |
| ☐ | `docker run --rm -v "$PWD/tests/samples/valid_code:/workspace" vbalidator:uat /workspace --quiet --no-strict` | Exit 0. |
| ☐ | `docker run --rm vbalidator:uat --version` *(if implemented)* or smoke against an inline file | Confirms the entrypoint is the installed `vbalidator`. |

---

## Section 11 — CI workflows on the actual GitHub run

Open <https://github.com/twobeass/VBAlidator/actions> and find the
latest run on the branch.

| ☐ | Workflow / job | Expected status |
|---|----------------|-----------------|
| ☐ | CI — `Lint (ruff)` | success |
| ☐ | CI — `Test (Py3.10..3.13 on ubuntu-latest)` | 4/4 success |
| ☐ | CI — `Test (Py3.10..3.13 on windows-latest)` | 4/4 success |
| ☐ | CI — `Test (Py3.10..3.13 on macos-latest)` | 4/4 success |
| ☐ | CI — `Rule docs in sync` | success |
| ☐ | CI — `CLI smoke test` | success |
| ☐ | Security — `pip-audit`, `bandit`, `CodeQL (Python)` | all 3 success |
| ☐ | Docker — `Build & push (multi-arch)` | success on PR (skipped push); pushes to GHCR after merge |
| ☐ | Docs — `Build site` | success |
| ☐ | PR Quality — `PR title is a Conventional Commit`, `commitlint`, `PR size label` | all success |

Failing checks in any of these are blockers for merge.

---

## Section 12 — Windows + Office (manual)

Only doable on a Windows machine with Office installed. Track in
`TODO.md` section A.

### 12a — Round-trip via VBE

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | `pip install pywin32` | pywin32 installed. |
| ☐ | Office → File ▶ Options ▶ Trust Center ▶ Trust Center Settings ▶ Macro Settings ▶ ☑ Trust access to the VBA project object model | Setting persisted. |
| ☐ | `vbalidator tests/samples/valid_code/valid_sample.bas --host excel --quiet --no-strict --roundtrip` | Exit 0. **No** `VBA_RT000` info, **no** `VBA_RT001` errors. |
| ☐ | `vbalidator tests/demo/BadModule.bas --host excel --quiet --roundtrip` | Exit 1. At least one `VBA_RT001` (severity `compile_verified`) — VBE itself rejects the module. |
| ☐ | Compare `result.errors` (static) vs the `VBA_RT001` set (dynamic) | Static catches a strict superset; if VBE finds something static missed, file an issue with the `false_positive.yml` template flipped. |

### 12b — Generate a real host model

| ☐ | Step | Expected |
|---|------|----------|
| ☐ | Open Excel → VBE → import `tools/VBA_Model_Exporter.bas` | Module loads without errors. |
| ☐ | Run `ExportReferences` macro | MsgBox shows the output path. `vba_references.json` contains `VBA`, `Excel`, `Office`, `stdole` at minimum. |
| ☐ | `python tools/generate_model.py /path/to/vba_references.json -o /tmp/excel-real.json -v` | Logs per-library progress. Output JSON has `globals`, `classes`, `enums` sections; `classes.Application.members` is non-empty. |
| ☐ | `vbalidator tests/samples/valid_code --model /tmp/excel-real.json --quiet --no-strict` | Exit 0. The richer model resolves at least as many identifiers as the bundled `--host excel` model. |
| ☐ | Repeat for Word / Access / Outlook | Each produces a host-specific model that `--model` can consume. |

### 12c — Check VBA_Model_Exporter on every host

| ☐ | Excel | `Application.VBE.ActiveVBProject` path used. References list contains `Excel`. |
| ☐ | Word | Falls through to the same path; `Word` reference present. |
| ☐ | Access | `Application.CurrentProject.Path` used for output; `Access` reference present. |
| ☐ | PowerPoint | `ThisDocument.VBProject` path; `PowerPoint` reference present. |
| ☐ | Outlook | VBE.ActiveVBProject path; `Outlook` reference present. |

---

## Section 13 — Sign-off

When every section above is ticked:

```bash
# Final sanity: everything green at the same time on a fresh clone
git switch claude/improve-vba-precompiler-3YcZp
pip install -e ".[dev,docs]"
pytest -ra && \
  ruff check src tests tools && \
  python tools/generate_rule_docs.py && \
  mkdocs build --strict
```

Expected: zero failing tests, zero ruff findings, zero doc-generator
diff, zero MkDocs warnings.

If yes → the branch is **ready to merge**.

If anything fails: capture the exact command, the full output, and the
`PrecheckResult.json()` dump if applicable; file an issue with the
`false_positive.yml` or `bug.yml` template.
