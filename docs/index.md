# VBAlidator

> **Premium VBA static analyser & compile-safety prechecker** — designed
> from the ground up to validate AI-generated VBA before it ever
> reaches a production workbook.

[![PyPI](https://img.shields.io/pypi/v/vbalidator)](https://pypi.org/project/vbalidator/)
[![CI](https://github.com/twobeass/VBAlidator/actions/workflows/ci.yml/badge.svg)](https://github.com/twobeass/VBAlidator/actions/workflows/ci.yml)
[![Docker](https://img.shields.io/badge/ghcr.io-twobeass%2Fvbalidator-blue)](https://github.com/twobeass/VBAlidator/pkgs/container/vbalidator)

VBAlidator parses `.bas` / `.cls` / `.frm` files, walks them through a
proper lexer / preprocessor / parser / analyzer pipeline, and emits a
**0–100 confidence score** plus a stable JSON v2 report you can drop
straight into any CI pipeline or LLM-VBA validator stack.

## Why VBAlidator

- **AI-grade prechecker.** Catches the typical hallucinations of
  LLM-VBA generators — wrong arity, missing `Set`, hallucinated method
  names, missing `PtrSafe`, type-mismatched literals — before VBE ever
  sees the code.
- **Optional dynamic verification.** With `--roundtrip` (Windows + Office)
  the static analyser is cross-checked against the actual VBE compiler.
- **Bundled host models.** `--host excel|word|access|outlook` ships the
  matching object model so you don't need to run a COM exporter first.
- **Stable rule IDs.** Every finding carries a `VBA<id>` you can pin in
  CI ignore lists. The catalogue at [Rules](rules/index.md) documents
  each one with a Failing example, a Compliant example and a fix hint.
- **CI-ready.** Confidence-Score, JSON v2 report, Docker image,
  Conventional-Commit-driven semantic releases on PyPI and GHCR.

## In one line

```bash
pip install vbalidator
vbalidator ./MyModules --host excel
```

## In one Python call

```python
from vbalidator import precheck

result = precheck("Module1.bas", host="excel")
if not result.compile_safe:
    for err in result.errors:
        print(f"{err['rule_id']}: {err['message']}")
print(f"score = {result.score}/100")
```

## Where to next

- [Quickstart](quickstart.md) — install, scan, ship in five minutes.
- [AI Pipeline Integration](ai-integration.md) — drop `precheck()`
  behind any LLM-VBA generator.
- [Rule catalogue](rules/index.md) — every emitted finding documented.
- [User Guide](Usage.md) — full CLI reference.
- [Architecture](Architecture.md) — how the pipeline works.
- [CI/CD](ci-cd.md) — release / Docker / docs / security workflows.
- [Roadmap](roadmap.md) — what's done, what's next.
