"""Phase 1.5 — block-terminator pairing (VBA110–VBA119).

Each block opener (`For`, `Do`, `While`, `With`, `Select Case`, block `If`)
must have its matching terminator. The individual block parsers now detect a
missing terminator; `parse_block` flags an inverse terminator with no opener.
"""
from __future__ import annotations

import pytest


def _ids(result) -> set[str]:
    return {e.get("rule_id") for e in result.errors}


CASES_MISSING = [
    ("VBA110", "Sub S()\n    For i = 1 To 10\n        Debug.Print i\nEnd Sub\n"),
    ("VBA111", "Sub S()\n    Do While x < 10\n        x = x + 1\nEnd Sub\n"),
    ("VBA112", "Sub S()\n    While x < 10\n        x = x + 1\nEnd Sub\n"),
    ("VBA113", "Sub S()\n    With foo\n        .Value = 1\nEnd Sub\n"),
    ("VBA114", "Sub S()\n    Select Case x\n        Case 1\n            Debug.Print 1\nEnd Sub\n"),
    ("VBA115", "Sub S()\n    If x > 0 Then\n        Debug.Print x\nEnd Sub\n"),
]


@pytest.mark.parametrize("rule_id, src", CASES_MISSING, ids=[c[0] for c in CASES_MISSING])
def test_missing_terminator_flagged(rule_id, src, run_source):
    result = run_source(src)
    assert rule_id in _ids(result), (
        f"Expected {rule_id} for missing terminator. Got: {result.messages!r}"
    )


CASES_INVERSE = [
    ("VBA116", "Sub S()\n    Debug.Print 1\n    Next i\nEnd Sub\n"),
    ("VBA117", "Sub S()\n    Debug.Print 1\n    Loop\nEnd Sub\n"),
    ("VBA118", "Sub S()\n    Debug.Print 1\n    Wend\nEnd Sub\n"),
]


@pytest.mark.parametrize("rule_id, src", CASES_INVERSE, ids=[c[0] for c in CASES_INVERSE])
def test_inverse_terminator_flagged(rule_id, src, run_source):
    result = run_source(src)
    assert rule_id in _ids(result), (
        f"Expected {rule_id} for stray terminator. Got: {result.messages!r}"
    )


def test_balanced_blocks_are_clean(run_source):
    src = (
        "Sub S()\n"
        "    Dim i As Long, x As Long\n"
        "    For i = 1 To 10\n"
        "        Do While x < 10\n"
        "            If x > 0 Then\n"
        "                x = x + 1\n"
        "            End If\n"
        "        Loop\n"
        "    Next i\n"
        "    While x > 0\n"
        "        x = x - 1\n"
        "    Wend\n"
        "    Select Case x\n"
        "        Case 0\n"
        "            Debug.Print 0\n"
        "    End Select\n"
        "End Sub\n"
    )
    result = run_source(src)
    terminator_ids = {f"VBA11{n}" for n in range(0, 10)}
    assert not (terminator_ids & {e.get("rule_id") for e in result.errors}), (
        f"Balanced blocks must not raise terminator errors. Got: {result.messages!r}"
    )
