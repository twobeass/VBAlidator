# GitHub Actions integration

Drop VBAlidator into the CI of any repository that holds VBA modules.
This page collects copy-pasteable workflow recipes from "smallest
possible gate" to "full PR-annotation + artifact + multi-host matrix".

The CLI exits **non-zero** on errors or below the score threshold, so
every example below is a single shell line; nothing GitHub-specific
needs to be wired up beyond the usual `actions/checkout`.

## TL;DR — the 5-line gate

Add this to `.github/workflows/vba-precheck.yml` in your VBA repo:

```yaml
name: VBA precheck
on: [push, pull_request]
jobs:
  precheck:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.12" }
      - run: pip install vbalidator
      - run: vbalidator ./vba --host excel --score-threshold 90
```

Replace `./vba` with the folder that holds your `.bas` / `.cls` /
`.frm` exports and pick the host that matches your codebase
(`excel` / `word` / `access` / `outlook` / `visio`). The job fails
when any module emits an error or the confidence score falls below
the threshold.

## Recipe 1 — Trigger only on VBA changes

Path filters keep the workflow off PRs that only touch READMEs or
unrelated code:

```yaml
name: VBA precheck
on:
  push:
    paths: ["**/*.bas", "**/*.cls", "**/*.frm"]
  pull_request:
    paths: ["**/*.bas", "**/*.cls", "**/*.frm"]

jobs:
  precheck:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.12", cache: "pip" }
      - run: pip install vbalidator
      - run: vbalidator ./vba --host excel --score-threshold 90
```

`cache: "pip"` shaves ~3 s off subsequent runs by reusing the wheel
download.

## Recipe 2 — Docker (no Python install needed)

If your repo never needs Python for anything else, the Docker image
is the leanest option:

```yaml
name: VBA precheck
on: [push, pull_request]
jobs:
  precheck:
    runs-on: ubuntu-latest
    container:
      image: ghcr.io/twobeass/vbalidator:latest
    steps:
      - uses: actions/checkout@v4
      - run: vbalidator . --host excel --score-threshold 90
```

The image is multi-arch (amd64 + arm64) and rebuilt on every
release, so `:latest` always tracks the most recent PyPI version.
Pin to a specific tag (`:v1.6.1`) if you want byte-stable runs.

## Recipe 3 — Upload the JSON v2 report as an artifact

Useful for postmortems and for downstream tools that consume the
canonical report:

```yaml
- run: vbalidator ./vba --host excel --output vba_report.json
- uses: actions/upload-artifact@v4
  if: always()           # also upload on failure
  with:
    name: vba-report
    path: vba_report.json
```

The report shape is documented under
[Usage → Issue shape (JSON v2)](Usage.md#issue-shape-json-v2) — every
finding carries a stable `rule_id` you can pin in dashboards.

## Recipe 4 — PR-level annotations (line-pinned errors)

GitHub renders `::error file=X,line=N::message` lines as inline
review comments. The JSON v2 report carries everything we need:

```yaml
- name: Run VBAlidator
  run: |
    vbalidator ./vba --host excel --quiet --output vba_report.json || true

- name: Annotate PR with findings
  run: |
    python -c '
    import json
    for i in json.load(open("vba_report.json"))["issues"]:
        sev = i["severity"]
        kind = "error" if sev in ("error", "compile_verified") else "warning"
        print(f"::{kind} file={i[\"file\"]},line={i[\"line\"]}::"
              f"[{i[\"rule_id\"]}] {i[\"message\"]}")
    '

- name: Re-fail on errors
  run: |
    python -c '
    import json, sys
    r = json.load(open("vba_report.json"))
    if not r["summary"]["compile_safe"]:
        sys.exit(1)
    '
```

The `|| true` after the first `vbalidator` call keeps the workflow
running so the annotation step can fire even when there are errors.
The last step rechecks the JSON and fails the job — annotations
without a real exit code don't block the merge.

## Recipe 5 — Matrix over multiple hosts

If your repo holds VBA for several Office hosts side-by-side
(say `vba/excel/` and `vba/word/`):

```yaml
jobs:
  precheck:
    runs-on: ubuntu-latest
    strategy:
      fail-fast: false
      matrix:
        include:
          - { host: excel, path: vba/excel }
          - { host: word,  path: vba/word  }
          - { host: access, path: vba/access }
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.12", cache: "pip" }
      - run: pip install vbalidator
      - run: vbalidator ${{ matrix.path }} --host ${{ matrix.host }} --score-threshold 90
```

`fail-fast: false` lets the other matrix entries run to completion
even when one host fails — you get one red square per problematic
codebase, not just the first.

## Recipe 6 — Step Summary for human review

`$GITHUB_STEP_SUMMARY` renders Markdown on the workflow summary page.
Drop a compact verdict so reviewers don't need to dig through logs:

```yaml
- name: Run VBAlidator
  run: vbalidator ./vba --host excel --output vba_report.json --quiet || true

- name: Step summary
  run: |
    python <<'PY' >> "$GITHUB_STEP_SUMMARY"
    import json
    r = json.load(open("vba_report.json"))
    s = r["summary"]
    icon = "✅" if s["compile_safe"] else "❌"
    print(f"# {icon} VBAlidator — score {s['score']} / 100")
    print()
    print(f"| Severity | Count |")
    print(f"|---|--:|")
    print(f"| errors   | {s['errors']} |")
    print(f"| warnings | {s['warnings']} |")
    print(f"| info     | {s['info']} |")
    print()
    if r["issues"]:
        print("## Findings")
        print("| File | Line | Rule | Message |")
        print("|---|--:|---|---|")
        for i in r["issues"][:50]:
            print(f"| {i['file']} | {i['line']} | {i['rule_id']} | {i['message']} |")
    PY

- name: Re-fail on errors
  run: |
    python -c 'import json,sys; sys.exit(0 if json.load(open("vba_report.json"))["summary"]["compile_safe"] else 1)'
```

## Recipe 7 — Local custom model alongside the bundled host

If your project has its own `.tlb` or references a class the bundled
host model doesn't ship (e.g. a private add-in), drop a
`vba_model.json` next to your VBA folder and VBAlidator auto-loads it:

```yaml
- run: vbalidator ./vba --host excel
  # ./vba/vba_model.json is auto-discovered — no flag needed.
```

Or be explicit:

```yaml
- run: vbalidator ./vba --host excel --model models/my_addin.json
```

See [Configuration → Custom models](Configuration.md#custom-models)
for the schema and [Configuration → Bundled host models](Configuration.md#bundled-host-models)
for the seven auto-layered COM-companion stubs (`scripting`,
`vbscript_regexp`, `wscript_shell`, `shell_application`, `mscomctl`,
`msforms`, plus the five Office hosts) — most repos don't need a
custom model because the auto-layering already covers
`Scripting.Dictionary`, `VBScript.RegExp`, MSForms UserForms, etc.

## Recipe 8 — Status badge in your VBA repo's README

```markdown
[![VBA precheck](https://github.com/<your-org>/<your-repo>/actions/workflows/vba-precheck.yml/badge.svg)](https://github.com/<your-org>/<your-repo>/actions/workflows/vba-precheck.yml)
```

## What VBAlidator catches (one-line recap)

Every emitted finding carries a `rule_id` you can pin in `jq` filters
or PR ignore-lists. The catalogue is at [Rules](rules/index.md). The
high-impact subset for AI-generated VBA:

- **VBA001** undefined identifiers, **VBA002** missing object members
- **VBA006** wrong argument count, **VBA007** ByRef type mismatches
- **VBA210** `Set` on scalar, **VBA211** missing `Set` on object assignment
- **VBA221–VBA224** Property Get/Let/Set arity & semantics
- **VBA230 / VBA231** non-constant Const initialisers
- **VBA300** missing `PtrSafe` on 64-bit Office Declares
- **VBA320** missing `Option Explicit` (warning)
- **VBA330** incomplete `Implements` of an interface
- **VBA340 / VBA341** `RaiseEvent` without matching `Event` declaration / wrong arity
- **VBA_LEX001 / VBA_LEX002** unrecognised characters & malformed date literals
- **VBA_RT001** *(optional)* errors caught by the actual VBE compiler via Office round-trip

## When you want VBE itself in the loop

The CLI's `--roundtrip` flag drives the actual VBE compiler over
Office COM for a second-opinion check, but that requires a Windows
runner with Office and pywin32 installed — GitHub-hosted
`windows-latest` doesn't ship Office. The setup is documented under
[CI/CD → Self-hosted Windows runner](ci-cd.md) for repos that have
the infrastructure.

For most repos the static analysis is enough; the round-trip is a
nice-to-have, not a requirement.
