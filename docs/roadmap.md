# Roadmap

VBAlidator was built in seven shippable PRs against the
["Premium-Prechecker für KI-generierten VBA-Code"](roadmap.md) plan.
This page tracks what's in, what's deferred, and what's next.

## Done

| PR | Phase | Highlights |
|----|-------|------------|
| #1 | 0 — Foundation | Single-line-If parser bug fix, lexer MISMATCH hardening, pytest migration, CI matrix (3 OS × 5 Py) |
| #2 | 1 — Control-flow | Real `For` / `Do` / `While` / `Select` body parsing, `ReDim` / `Erase` validation, lexer legacy-token support (`$`/`%`/`@` suffixes, `[bracket]` foreign names), 80-entry `std_model` backfill |
| #3 | 2.1 / 2.2 / 2.3 | Label registry + jump validation, Set-vs-Let on bare-identifier LHS, Property Get/Let/Set arity & semantic checks |
| #4 | 2.4 / 2.5 / 2.7 / 2.8 / 2.9 | Operator-type literal check, Const-expression validation, structural date-literal parser, fixed-length-string scope check, `DefInt`/`DefStr`/… implicit typing |
| #5 | 3.1 / 3.2 / 3.3 / 3.4 / 3.6 | `Implements`-interface contract, `Event` / `RaiseEvent` / `WithEvents`, `PtrSafe` requirement, Enum-member uniqueness, `Option Explicit` warning |
| #6 | 4.1 / 4.2 / 4.3 / 4.4 | Bundled host models (Excel/Word/Access/Outlook), 0–100 confidence score, JSON v2 schema, `from vbalidator import precheck` Python API |
| #7 | 4.5 / 4.6 | Round-trip via Office COM, rule registry + auto-generated catalogue (`docs/rules/`) |
| #8 | 5 | Maximum CI/CD: semantic-release → PyPI Trusted Publisher, multi-arch Docker → GHCR, security pipeline, MkDocs → GitHub Pages, repo hygiene |

## Deferred — high-value, queued for next iteration

- **P3.5 — Module-Level vs. Procedure-Level Statement Placement.**
  Reject e.g. `Type … End Type` inside a `Sub`. Needs a parser-level
  refactor to track the active container.
- **P2.6 — UDT / Class member-chain depth.** Today member chains
  validate the first hop fully and degrade to permissive Variant past
  the second `.`. Closing this requires a proper type system on the
  walker.
- **VBA350 — `End` placement in Function-returns.** AI generators sometimes
  emit `End Sub` to close a Function. Easy rule, just needs a fixture.

## Deferred — research / nice-to-have

- **Auto-fix engine.** The catalogue's `fix_hint` field is the seed.
  Most useful for: missing `Set`, missing `PtrSafe`, wrong Property
  arity, Levenshtein-suggesting typo'd built-in names.
- **VS Code extension.** LSP server wrapping `precheck()` for
  inline diagnostics on `.bas` / `.cls` / `.frm` files. Lives in a
  separate repo; the release workflow can dispatch a
  `repository_dispatch` event to trigger its build.
- **Web demo.** Streamlit / Gradio app — paste code, get score.

## Scoring philosophy

A score reaches 100 only when zero issues are present. Each severity
subtracts a fixed weight from a starting 100:

```text
score = max(0, 100 − Σ(weight × count))
weights: error=20  warning=3  info=1  compile_verified=30
```

The intent is: a single hard error always drops the score below the
default 90 % CI gate. Style warnings (e.g. missing `Option Explicit`)
nudge the score down by 3 each so a clean module landing at 100
genuinely means *both* compile-safe *and* style-clean.

## Versioning

Conventional-Commit-driven semver via python-semantic-release:

| Commit prefix | Version bump |
|---------------|--------------|
| `feat:` | minor |
| `fix:`, `perf:`, `refactor:` | patch |
| Body footer `BREAKING CHANGE:` | major |
| `docs:`, `test:`, `chore:`, `ci:`, `build:`, `style:` | none |

Stable rule IDs are part of the public contract — they never change
across releases. New rules get new IDs; deprecations are documented
in the changelog before removal.
