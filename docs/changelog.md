# Changelog

This page is auto-managed by `python-semantic-release` from
Conventional Commits. The same content also lives at
[`CHANGELOG.md`](https://github.com/twobeass/VBAlidator/blob/main/CHANGELOG.md)
in the repository root.

Until the first semantic-release tag lands, the section below tracks
the major milestones manually.

## Unreleased

### feat
- **CI/CD pipeline (Phase 5).** Maximum-config GitHub Actions setup:
  semantic-release → PyPI (OIDC Trusted Publisher), multi-arch Docker →
  GHCR, MkDocs → GitHub Pages, security workflow (pip-audit / bandit /
  CodeQL), nightly Windows VBE round-trip, weekly drift snapshot.
- **Documentation overhaul.** New Quickstart, AI Pipeline Integration
  guide, expanded Architecture / Usage / Configuration, full Rule
  catalogue (35 entries) at `docs/rules/`.

## 0.1.x — Phase 0–4 milestones

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
