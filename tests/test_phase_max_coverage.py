"""Maximum-coverage additions: VBA151, VBA270, VBA370."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_code_after_end_sub(run_source):
    r = run_source("Sub S()\nEnd Sub: x = 1\n")
    assert "VBA151" in _ids(r)


def test_comment_after_end_sub_ok(run_source):
    r = run_source("Sub S()\nEnd Sub   ' all done\n")
    assert "VBA151" not in _ids(r)


def test_plain_end_sub_ok(run_source):
    r = run_source("Sub S()\nEnd Sub\n")
    assert "VBA151" not in _ids(r)


def test_type_suffix_mismatch(run_source):
    r = run_source("Sub S()\n    Dim count% As Long\nEnd Sub\n")
    assert "VBA270" in _ids(r)


def test_type_suffix_match_ok(run_source):
    r = run_source("Sub S()\n    Dim count% As Integer\nEnd Sub\n")
    assert "VBA270" not in _ids(r)


def test_type_suffix_no_as_ok(run_source):
    r = run_source("Sub S()\n    Dim count%\nEnd Sub\n")
    assert "VBA270" not in _ids(r)


def test_public_const_in_class(run_source):
    r = run_source('Attribute VB_Name = "C"\nPublic Const MAX As Long = 10\n', module_type="Class")
    assert "VBA370" in _ids(r)


def test_public_array_in_class(run_source):
    r = run_source('Attribute VB_Name = "C"\nPublic items() As Long\n', module_type="Class")
    assert "VBA370" in _ids(r)


def test_private_const_in_class_ok(run_source):
    r = run_source('Attribute VB_Name = "C"\nPrivate Const MAX As Long = 10\n', module_type="Class")
    assert "VBA370" not in _ids(r)


def test_public_const_in_standard_module_ok(run_source):
    r = run_source('Attribute VB_Name = "M"\nPublic Const MAX As Long = 10\n', module_type="Module")
    assert "VBA370" not in _ids(r)
