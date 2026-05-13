# Changelog

The authoritative changelog for **v1.0.0 onwards** lives at
<https://github.com/twobeass/VBAlidator/releases> — regenerated on
every `main` push by `python-semantic-release` from the Conventional
Commit history.

This page keeps the pre-release Phase 0–5 milestone history so the
project's lineage stays visible.

## v1.x highlights (full detail in GitHub Releases)

- **v1.4.x** — auto-layer 4 COM library stubs (`scripting`,
  `vbscript_regexp`, `wscript_shell`, `shell_application`);
  default-property `Item` resolution closes the last analyzer FP in
  `awesome_vba` (VBAlidator's FP surface is now zero).
- **v1.3.x** — Iter-6 + Iter-7 false-positive sweep:
  `As Any` / `As Any()` ByRef sentinels, MSComCtl host model with
  `.frm` auto-layering, MSForms 2.0 host model with namespace
  auto-layering, suppressing Sub-style implicit-call inside
  expressions, Enum↔Long ByRef bidirectional compat (drops 62
  false-positives across stdLambda alone), `Erase` / `ReDim` member-
  chain support, `error` as identifier, lexer trailing-whitespace
  before `_` line continuation, Enum members as `EnumItem` constants
  (valid as `Const X = MyEnum.Member` RHS). Drives awesome_vba from
  203 → 4 hard errors, all four remaining are genuine upstream bugs.
- **v1.1.x — v1.2.x** — full-fidelity Excel/Word/Access/Visio host
  models (1–3 MB each, ~5000 globals/classes total) replace the
  earlier hand-curated stubs; awesome_vba host-aware regression
  runner; `stdError.cls` fixture completion + parser hang fix on
  standalone `End` inside If/With/Select blocks; semantic-release
  pipeline activated.

## Pre-release: Phase 0–5 milestones

### Phase 4.5–4.6 (PR #7)
- `--roundtrip` Office-COM cross-check (Windows + pywin32).
- `src/rules.py` single source of truth for every emitted finding.
- Auto-generated rule catalogue at `docs/rules/`.
- CI gate: `tools/generate_rule_docs.py` must produce a clean diff.

### Phase 4.1–4.4 (PR #6)
- `from vbalidator import precheck` public API.
- 0–100 confidence score with severity weights (error=20, warning=3,
  info=1, compile_verified=30).
- JSON v2 report schema with stable rule IDs.
- Bundled host models for Excel / Word / Access / Outlook.

### Phase 3.1–3.6 (PR #5)
- `Implements`-interface contract (VBA330).
- `Event` + `RaiseEvent` (VBA340 / VBA341).
- `PtrSafe` requirement on 64-bit Office (VBA300).
- Enum-member uniqueness (VBA310).
- `Option Explicit` warning (VBA320).
- Preprocessor case-insensitive define lookup; modern Office defaults.

### Phase 2.4–2.9 (PR #4)
- Operator-type literal check (VBA240).
- Const-expression validation (VBA230, VBA231).
- Structural date-literal parser (VBA_LEX002).
- Fixed-length-string scope check (VBA250).
- `DefInt`/`DefStr`/… implicit typing.

### Phase 2.1–2.3 (PR #3)
- Label registry & jump validation (VBA201).
- Set-vs-Let on bare-identifier LHS (VBA210, VBA211).
- Property Get/Let/Set arity & semantic checks (VBA221–VBA224).

### Phase 1 (PR #2)
- Real control-flow body parsing (For / Do / While / Select).
- ReDim / Erase validation (VBA101–VBA106).
- Lexer legacy-token support (`$`/`%`/`@` suffixes, `[A1]` brackets).
- `std_model.json` backfill (~80 missing built-ins).

### Phase 0 (PR #1)
- Single-line-If parser bug fix.
- Lexer MISMATCH hardening (VBA_LEX001).
- Pytest migration, CI matrix, awesome-VBA regression baseline.
