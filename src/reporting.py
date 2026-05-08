"""Issue normalization + JSON v2 report builder.

Stable rule IDs are assigned to every analyzer / lexer / parser finding
so CI pipelines can build ignore lists without depending on free-form
message text. Legacy findings (the original analyzer rules added before
the rule-ID convention) are mapped from message patterns at report time;
new rules added in Phase 1+ supply their own rule_id directly.
"""
from __future__ import annotations

import re
from typing import Iterable

# Regex → (rule_id, severity, category) for legacy issue messages.
# Order matters: first match wins. Be precise.
_LEGACY_PATTERNS: list[tuple[re.Pattern, str, str, str]] = [
    (re.compile(r"^Unreachable code"), "VBA009", "warning", "control_flow"),
    (re.compile(r"^Exit .* not allowed"), "VBA008", "error", "control_flow"),
    (re.compile(r"^Undefined identifier"), "VBA001", "error", "name_resolution"),
    (re.compile(r"^Member .* not found"), "VBA002", "error", "member_access"),
    (re.compile(r"^Duplicate declaration"), "VBA003", "error", "declaration"),
    (re.compile(r"^Invalid or unexpected . reference"), "VBA004", "error", "member_access"),
    (re.compile(r"^Expected Array or Procedure"), "VBA005", "error", "type"),
    (re.compile(r"^Argument count mismatch"), "VBA006", "error", "signature"),
    (re.compile(r"^ByRef argument type mismatch"), "VBA007", "error", "signature"),
    (re.compile(r"^Syntax Error"), "VBA010", "error", "syntax"),
]


def _infer_rule(message: str) -> tuple[str, str, str]:
    """Return (rule_id, severity, category) for a legacy free-form message."""
    for pat, rid, sev, cat in _LEGACY_PATTERNS:
        if pat.match(message or ""):
            return rid, sev, cat
    return "VBA000", "error", "unknown"


_CATEGORY_BY_RULE_PREFIX: dict[str, str] = {
    # Phase IDs follow the roadmap numbering (VBA1xx Phase 1, VBA2xx
    # Phase 2, VBA3xx Phase 3, VBA_LEX… lexer-level).
    "VBA10": "control_flow",     # ReDim / Erase
    "VBA20": "jump",             # GoTo / Resume / On Error
    "VBA21": "assignment",       # Set / Let
    "VBA22": "property",         # Property arity / semantics
    "VBA23": "const_expression",
    "VBA24": "operator_type",
    "VBA25": "fixed_string",
    "VBA30": "platform",         # PtrSafe
    "VBA31": "enum",
    "VBA32": "style",            # Option Explicit
    "VBA33": "interface",        # Implements
    "VBA34": "events",           # RaiseEvent
    "VBA_LEX": "lexer",
}


def _category_for(rule_id: str) -> str:
    if rule_id is None:
        return "unknown"
    for prefix, cat in _CATEGORY_BY_RULE_PREFIX.items():
        if rule_id.startswith(prefix):
            return cat
    return "unknown"


def normalize_issue(issue: dict) -> dict:
    """Return a copy of `issue` with rule_id, severity, category and a
    canonical key set populated. Idempotent.
    """
    out = dict(issue)
    rid = out.get("rule_id")
    if not rid:
        rid, sev, cat = _infer_rule(out.get("message", ""))
        out["rule_id"] = rid
        out.setdefault("severity", sev)
        out.setdefault("category", cat)
    else:
        out.setdefault("severity", "error")
        out.setdefault("category", _category_for(rid))
    out.setdefault("file", "")
    out.setdefault("line", 0)
    out.setdefault("column", 0)
    return out


def normalize_issues(issues: Iterable[dict]) -> list[dict]:
    return [normalize_issue(i) for i in issues]


def build_report_v2(
    issues: Iterable[dict],
    files_scanned: int,
    score: int,
    compile_safe: bool,
    score_breakdown: dict | None = None,
) -> dict:
    """Return the canonical JSON v2 report for CI / API consumers."""
    norm = normalize_issues(issues)
    counts = {"error": 0, "warning": 0, "info": 0}
    for i in norm:
        sev = i.get("severity", "error")
        if sev not in counts:
            counts[sev] = 0
        counts[sev] += 1

    # Group issues per file for quick CI navigation.
    per_file: dict[str, list[dict]] = {}
    for i in norm:
        per_file.setdefault(i.get("file", ""), []).append(i)

    files_payload = []
    for fname, items in sorted(per_file.items()):
        items_sorted = sorted(items, key=lambda x: (x.get("line", 0), x.get("column", 0)))
        files_payload.append({"path": fname, "issues": items_sorted})

    return {
        "version": "2.0",
        "summary": {
            "score": score,
            "compile_safe": compile_safe,
            "errors": counts["error"],
            "warnings": counts["warning"],
            "info": counts["info"],
            "files_scanned": files_scanned,
            "issues_total": len(norm),
        },
        "score_breakdown": score_breakdown or {},
        "files": files_payload,
        "issues": norm,  # flat list, useful for `jq '.issues[]'`
    }
