"""Sample-driven tests for `tests/samples/compile_errors/<category>/*.bas`.

For each sample file we expect at least one analyzer error whose message
relates to the category name (loose substring match of category keywords).
This makes it easy to add new fixtures: drop a `.bas` file into the right
category folder and it is picked up automatically.
"""
from __future__ import annotations

from pathlib import Path

import pytest


CATEGORY_KEYWORDS: dict[str, tuple[str, ...]] = {
    "undefined_identifier": ("undefined", "not found"),
    "duplicate_declaration": ("duplicate",),
    "member_access": ("member", "without with", "not found"),
    "type_mismatch": ("expected", "type", "mismatch"),
    "argument_mismatch": ("argument count", "expected at"),
    "byref_mismatch": ("byref", "argument type mismatch"),
    "unreachable_code": ("unreachable",),
    "syntax_errors": ("syntax", "unexpected"),
    "redim_target": ("redim", "undefined"),
    "erase_target": ("erase",),
    "jump_target": ("not a label", "label"),
    "set_vs_let": ("set", "object assignment"),
    "property_arity": ("property", "must have"),
}


def _collect_samples(samples_root: Path) -> list[tuple[str, Path]]:
    out: list[tuple[str, Path]] = []
    compile_errors_root = samples_root / "compile_errors"
    if not compile_errors_root.is_dir():
        return out
    for category_dir in sorted(compile_errors_root.iterdir()):
        if not category_dir.is_dir():
            continue
        for sample in sorted(category_dir.iterdir()):
            if sample.suffix.lower() in (".bas", ".cls", ".frm"):
                out.append((category_dir.name, sample))
    return out


def _id_for(case: tuple[str, Path]) -> str:
    category, path = case
    return f"{category}/{path.name}"


SAMPLES = _collect_samples(Path(__file__).resolve().parent / "samples")


@pytest.mark.parametrize("category, sample_path", SAMPLES, ids=[_id_for(c) for c in SAMPLES])
def test_compile_error_sample_emits_relevant_error(category, sample_path, run_files):
    result = run_files([sample_path])

    assert result.errors, (
        f"Expected at least one compile error for {sample_path.relative_to(sample_path.parents[3])}, "
        f"but analyzer returned an empty error list."
    )

    keywords = CATEGORY_KEYWORDS.get(category, ())
    if not keywords:
        return  # Category without strict keyword expectations still passes if it errored at all.

    assert any(
        result.has_message_containing(kw) for kw in keywords
    ), (
        f"None of the messages for category '{category}' contained any of "
        f"{keywords}. Got: {result.messages!r}"
    )
