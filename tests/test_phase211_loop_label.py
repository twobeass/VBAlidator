"""Phase 2.11 — loop / label context (VBA170-VBA175)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_exit_for_outside_loop(run_source):
    r = run_source("Sub S()\n    Exit For\nEnd Sub\n")
    assert "VBA170" in _ids(r)


def test_exit_do_outside_loop(run_source):
    r = run_source("Sub S()\n    Exit Do\nEnd Sub\n")
    assert "VBA171" in _ids(r)


def test_exit_for_inside_loop_ok(run_source):
    r = run_source("Sub S()\n    Dim i As Long\n    For i = 1 To 3\n        Exit For\n    Next i\nEnd Sub\n")
    assert "VBA170" not in _ids(r)


def test_exit_do_inside_loop_ok(run_source):
    r = run_source("Sub S()\n    Do\n        Exit Do\n    Loop\nEnd Sub\n")
    assert "VBA171" not in _ids(r)


def test_exit_do_inside_while_wend_flagged(run_source):
    # Exit Do is not valid inside While...Wend.
    r = run_source("Sub S()\n    Dim x As Long\n    While x < 3\n        Exit Do\n    Wend\nEnd Sub\n")
    assert "VBA171" in _ids(r)


def test_next_variable_mismatch(run_source):
    r = run_source("Sub S()\n    Dim i As Long\n    For i = 1 To 3\n        Debug.Print i\n    Next j\nEnd Sub\n")
    assert "VBA172" in _ids(r)


def test_next_variable_match_ok(run_source):
    r = run_source("Sub S()\n    Dim i As Long\n    For i = 1 To 3\n        Debug.Print i\n    Next i\nEnd Sub\n")
    assert "VBA172" not in _ids(r)


def test_for_counter_reused_nested(run_source):
    r = run_source(
        "Sub S()\n    Dim i As Long\n    For i = 1 To 3\n        For i = 1 To 3\n        Next i\n    Next i\nEnd Sub\n"
    )
    assert "VBA173" in _ids(r)


def test_for_each_scalar_counter(run_source):
    r = run_source(
        "Sub S()\n    Dim i As Long\n    Dim coll As Collection\n    For Each i In coll\n    Next i\nEnd Sub\n"
    )
    assert "VBA174" in _ids(r)


def test_for_each_variant_counter_ok(run_source):
    r = run_source(
        "Sub S()\n    Dim v As Variant\n    Dim coll As Collection\n    For Each v In coll\n    Next v\nEnd Sub\n"
    )
    assert "VBA174" not in _ids(r)


def test_duplicate_label(run_source):
    r = run_source("Sub S()\nDone:\n    Exit Sub\nDone:\nEnd Sub\n")
    assert "VBA175" in _ids(r)


def test_single_label_ok(run_source):
    r = run_source("Sub S()\nDone:\n    Exit Sub\nEnd Sub\n")
    assert "VBA175" not in _ids(r)
