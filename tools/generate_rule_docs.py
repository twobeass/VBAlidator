#!/usr/bin/env python3
"""Regenerate `docs/rules/` from the central rule registry.

Run this whenever a new rule is added to `src/rules.py`. The script
produces:

- `docs/rules/index.md` — sortable table of every rule
- `docs/rules/<rule_id>.md` — one page per rule with What / Why / fail
  example / ok example / fix hint

The output is deterministic so `git diff` after running cleanly shows
exactly which rules changed.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Make `src` importable when run from the repo root.
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.rules import all_rules, Rule  # noqa: E402

DOCS_DIR = ROOT / "docs" / "rules"


def _severity_badge(sev: str) -> str:
    return {
        "error": "🔴 error",
        "warning": "🟡 warning",
        "info": "🔵 info",
        "compile_verified": "🔴 compile_verified",
    }.get(sev, sev)


def _rule_page(rule: Rule) -> str:
    parts = [
        f"# {rule.rule_id} — {rule.title}",
        "",
        f"**Severity:** {_severity_badge(rule.severity)}    "
        f"**Category:** `{rule.category}`    "
        f"**Phase:** {rule.phase or '—'}",
        "",
        "## Description",
        "",
        rule.description,
        "",
    ]
    if rule.fail_example:
        parts += [
            "## Failing example",
            "",
            "```vb",
            rule.fail_example,
            "```",
            "",
        ]
    if rule.ok_example:
        parts += [
            "## Compliant example",
            "",
            "```vb",
            rule.ok_example,
            "```",
            "",
        ]
    if rule.fix_hint:
        parts += [
            "## How to fix",
            "",
            rule.fix_hint,
            "",
        ]
    parts += [
        "---",
        "",
        # Absolute GitHub URL — keeps the link valid both when the page
        # is rendered through MkDocs (where `../../src/...` would land
        # outside the docs tree and break --strict) and on github.com.
        f"_Source: [src/rules.py](https://github.com/twobeass/VBAlidator/blob/main/src/rules.py) — entry `{rule.rule_id}`._",
        "",
    ]
    return "\n".join(parts)


def _index_page(rules: list[Rule]) -> str:
    lines = [
        "# VBAlidator rule catalogue",
        "",
        "Stable rule IDs emitted by VBAlidator. Each row links to the rule's"
        " detail page. Use the `rule_id` to silence specific findings in"
        " your CI ignore list — the IDs do not change between releases.",
        "",
        "| Rule | Severity | Category | Phase | Title |",
        "| --- | --- | --- | --- | --- |",
    ]
    for r in rules:
        lines.append(
            f"| [`{r.rule_id}`]({r.rule_id}.md) "
            f"| {_severity_badge(r.severity)} "
            f"| `{r.category}` "
            f"| {r.phase or '—'} "
            f"| {r.title} |"
        )
    lines += [
        "",
        f"*{len(rules)} rules registered.* "
        "Generated from `src/rules.py` via `python tools/generate_rule_docs.py`.",
        "",
    ]
    return "\n".join(lines)


def main() -> None:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    rules = all_rules()

    # Per-rule pages.
    # The catalogue contains emoji severity badges (🔴 / 🟡 / 🔵) so we
    # always read & write UTF-8 explicitly. Path.read_text / write_text
    # fall back to the platform default (cp1252 on Windows) without an
    # explicit encoding, which would crash on the emoji glyphs.
    written = 0
    expected_files = {"index.md"}
    for rule in rules:
        page = _rule_page(rule)
        path = DOCS_DIR / f"{rule.rule_id}.md"
        existing = path.read_text(encoding="utf-8") if path.exists() else None
        if existing != page:
            path.write_text(page, encoding="utf-8")
            written += 1
        expected_files.add(path.name)

    # Index
    index_path = DOCS_DIR / "index.md"
    index_text = _index_page(rules)
    existing_index = index_path.read_text(encoding="utf-8") if index_path.exists() else None
    if existing_index != index_text:
        index_path.write_text(index_text, encoding="utf-8")
        written += 1

    # Prune stale rule pages.
    for existing in DOCS_DIR.iterdir():
        if existing.is_file() and existing.name not in expected_files:
            existing.unlink()
            written += 1

    print(f"docs/rules/: {len(rules)} rules, {written} files updated.")


if __name__ == "__main__":
    main()
