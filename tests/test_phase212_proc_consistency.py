"""Phase 2.12 — procedure / property consistency (VBA150, VBA152, VBA153)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_duplicate_sub(run_source):
    r = run_source("Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\n")
    assert "VBA150" in _ids(r)


def test_distinct_subs_ok(run_source):
    r = run_source("Sub Foo()\nEnd Sub\nSub Bar()\nEnd Sub\n")
    assert "VBA150" not in _ids(r)


def test_property_get_let_share_name_ok(run_source):
    r = run_source(
        "Property Get Foo() As Long\nEnd Property\n"
        "Property Let Foo(ByVal v As Long)\nEnd Property\n"
    )
    assert "VBA150" not in _ids(r)


def test_property_get_let_type_mismatch(run_source):
    r = run_source(
        "Property Get Foo() As Long\nEnd Property\n"
        "Property Let Foo(ByVal v As String)\nEnd Property\n"
    )
    assert "VBA152" in _ids(r)


def test_property_get_let_type_match_ok(run_source):
    r = run_source(
        "Property Get Foo() As Long\nEnd Property\n"
        "Property Let Foo(ByVal v As Long)\nEnd Property\n"
    )
    assert "VBA152" not in _ids(r)


def test_implements_signature_mismatch(run_source, tmp_path):
    # Two modules: interface class + implementing class.
    from tests.conftest import run_pipeline_on_files
    iface = tmp_path / "IShape.cls"
    iface.write_text('Attribute VB_Name = "IShape"\nPublic Sub Draw(ByVal scale As Double)\nEnd Sub\n')
    impl = tmp_path / "Square.cls"
    impl.write_text('Attribute VB_Name = "Square"\nImplements IShape\nPrivate Sub IShape_Draw()\nEnd Sub\n')
    result = run_pipeline_on_files([iface, impl])
    assert "VBA153" in {e.get("rule_id") for e in result.errors}


def test_implements_signature_match_ok(run_source, tmp_path):
    from tests.conftest import run_pipeline_on_files
    iface = tmp_path / "IShape.cls"
    iface.write_text('Attribute VB_Name = "IShape"\nPublic Sub Draw(ByVal scale As Double)\nEnd Sub\n')
    impl = tmp_path / "Square.cls"
    impl.write_text('Attribute VB_Name = "Square"\nImplements IShape\nPrivate Sub IShape_Draw(ByVal scale As Double)\nEnd Sub\n')
    result = run_pipeline_on_files([iface, impl])
    ids = {e.get("rule_id") for e in result.errors}
    assert "VBA153" not in ids
