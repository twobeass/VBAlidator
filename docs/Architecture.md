# Architecture

VBAlidator is a layered pipeline. Each stage is a pure function over
the previous stage's output, which keeps the analyser deterministic
and trivially testable.

```text
       ┌───────────┐    ┌──────────────┐    ┌──────────┐    ┌──────────┐
src ─▶ │   Lexer   │ ─▶ │ Preprocessor │ ─▶ │  Parser  │ ─▶ │ Analyzer │ ─▶ Issues
       └───────────┘    └──────────────┘    └──────────┘    └──────────┘
                                                                  │
                                                                  ▼
                                                       ┌──────────────────┐
                                                       │  Scoring + JSON  │
                                                       └──────────────────┘
                                                                  │
                                       (optional, Windows + Office)│
                                                                  ▼
                                                       ┌──────────────────┐
                                                       │  VBE Round-trip  │
                                                       └──────────────────┘
```

## Lexer (`src/lexer.py`)

Regex-table tokenizer producing a `Token` stream. Recognises:

- comments, strings, line continuations (` _`)
- numeric literals with **legacy type-suffixes** (`100&` Long, `1.5#`
  Double, `50023612.1134@` Currency)
- identifiers with `$`/`%`/`@` suffixes (`Mid$`, `i%`, `c@`)
- bracket-quoted foreign names (`[A1]`, `[Sheet1!A1]`)
- `#…#` date literals — validated structurally (rule `VBA_LEX002`)
- preprocessor directives (`#If`, `#Const`, `#End`)

Unrecognised characters go through a hardened MISMATCH path that
records a `LexerError` (rule `VBA_LEX001`) instead of being silently
dropped.

## Preprocessor (`src/preprocessor.py`)

Stack-based scope evaluation of `#If` / `#ElseIf` / `#Else` /
`#End If` and `#Const` directives. Symbol lookup is **case-insensitive**
(VBA semantic), so `#If Vba7 Then` and `#If VBA7 Then` are equivalent.

Default constants reflect a modern Microsoft 365 host:

```python
{"VBA7": True, "WIN64": True, "WIN32": False, "WIN16": False, "MAC": False}
```

Override via `precheck(defines={…})` or `--define`.

## Parser (`src/parser.py`)

Recursive-descent parser producing an AST. Major node types:

| Node | Purpose |
|------|---------|
| `ModuleNode` | Top-level container (`.bas` / `.cls` / `.frm`). Tracks `options{explicit, compare, base, private_module}`, `def_type_map`, `implements[]`. |
| `ProcedureNode` | Sub / Function / Property Get/Let/Set / Event / Declare. Carries `is_declare`, `is_ptrsafe`, `args[]`, `body[]`. |
| `VariableNode` | `Dim`/`Const`/parameter; `is_const` flag distinguishes constants. |
| `TypeNode` | UDT / Enum. |
| `IfNode` | Multi- and single-line `If`/`ElseIf`/`Else`. |
| `ForNode` / `DoNode` / `SelectNode` / `CaseClauseNode` / `WithNode` | Real control-flow nodes — bodies are recursively walked, not skipped. |
| `RedimNode` / `EraseNode` | Array-resize/erase with target validation. |
| `StatementNode` | Catch-all token sequence for primitive statements. |

Parser-side errors (`VBA010` Syntax Error, missing `Then`, stray
block terminators) accumulate on `parser.errors` and merge into the
analyser's issue list.

## Analyzer (`src/analyzer.py`)

Two passes over every module:

- **Pass 1 — Discovery.** Build the `Global` and `Module` symbol
  tables. Public symbols, classes, library references, enum members,
  module-level Const/Dim declarations land in `Global`. UDTs and
  module-private declarations land in `Module`.
- **Pass 2 — Verification.** Per procedure:
    1. Build `Procedure` scope and seed it with parameters.
    2. Run a **Pass 1.5** label sweep across the procedure body
       (recursively into every nested control-flow node) to populate
       the jump-target registry consumed by `VBA201`.
    3. Walk every statement node, dispatching on type:
       - StatementNode → identifier resolution, signature checks,
         dotted member lookup.
       - IfNode / WithNode / ForNode / DoNode / SelectNode → analyse
         condition tokens then recurse into bodies.
       - RedimNode / EraseNode → target-existence and array-typedness.

A small set of validators are layered on top of pass 2:

| Validator | Phase | Rule IDs |
|-----------|-------|----------|
| `_validate_jump_target` | 2.1 | VBA201 |
| `_validate_set_vs_let` | 2.2 | VBA210, VBA211 |
| `_validate_property_arity` | 2.3 | VBA221–VBA224 |
| `_validate_operator_types` | 2.4 | VBA240 |
| `_validate_const_expression` | 2.5 | VBA230, VBA231 |
| `_validate_ptrsafe_declares` | 3.3 | VBA300 |
| `_validate_enum_uniqueness` | 3.4 | VBA310 |
| `_validate_option_explicit` | 3.6 | VBA320 |
| `_validate_implements` | 3.1 | VBA330 |
| `_validate_raise_event` | 3.2 | VBA340, VBA341 |

Each validator is independent and pure-ish — easy to disable, profile,
or backport into a custom subclass.

## Symbol resolution

`SymbolTable` is a parent-pointer chain with case-insensitive lookup
that **normalises identifier suffixes**. `Mid$`, `Mid`, `[Mid]` all
resolve to the same standard global. The chain is:

```text
Procedure ─▶ Module ─▶ Global ─▶ (built-in std_model + host model)
```

When a member chain `a.b.c.d` is walked, each step consults the loaded
object model first (Excel/Word/Access/Outlook), then UDT members, then
falls through to a permissive Variant. Forms (`.frm`) treat unknown
identifiers as implicit Controls so user-form fields don't trip the
analyser.

## Reporting (`src/reporting.py`)

The analyser emits raw issue dicts. `normalize_issues` decorates them
with stable `rule_id`, `severity`, `category` — inferring from message
patterns for the legacy (Phase-0) rules. `build_report_v2` produces
the canonical JSON schema:

```json
{
  "version": "2.0",
  "summary": {
    "score": 87, "compile_safe": false,
    "errors": 1, "warnings": 2, "info": 0,
    "files_scanned": 5, "issues_total": 3
  },
  "score_breakdown": { "starting": 100, "penalty_total": 13, "by_severity": {...} },
  "files": [{ "path": "Module1.bas", "issues": [...] }],
  "issues": [...]
}
```

## Scoring (`src/scoring.py`)

```text
score = max(0, 100 − Σ(severity_weight × count))

weights:  error=20  warning=3  info=1  compile_verified=30
```

`compile_safe` is True iff zero blocking findings (errors +
compile_verified). `coverage_uncertain=True` caps the score at 90 to
flag unresolved external library references.

## Round-trip (`src/roundtrip.py`)

Optional Phase-4.5 dynamic verification. On Windows + Office +
pywin32:

1. Inject the source into a temporary `.xlsm`/`.docm`.
2. Spin up the Office host with `Visible=False`.
3. Call `VBProject.Compile` and capture VBE's verdict.
4. Reflect any compile error back as a `compile_verified` issue.

On other platforms the module exposes a clean `RoundtripUnavailable`
exception and the CLI emits a single `VBA_RT000` info notice instead
of crashing.

## Rule registry (`src/rules.py`)

Single source of truth: every rule_id has a `Rule` dataclass with
title, severity, category, phase, description, fail / ok examples
and a fix hint. The `tools/generate_rule_docs.py` script consumes the
registry to regenerate `docs/rules/<id>.md` and `docs/rules/index.md`.
A CI job (`docs` in `ci.yml`) refuses to merge any PR where the
generator would produce a diff, so docs and code never drift.

## Public API (`src/api.py`)

Two entry points:

- `precheck(source, host=…, model_path=…, defines=…, strict=…, roundtrip=…)`
- `precheck_source(code, name=…, host=…, …)` — convenience for inline
  source strings.

Both are thin orchestration shims over the pipeline so the CLI and
the Python API never drift.
