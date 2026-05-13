<!--
Thanks for contributing to VBAlidator!

PR titles must follow Conventional Commits (enforced by CI):
  feat(precompiler): add VBA350 rule
  fix(parser): handle nested With blocks
  docs(rules): clarify VBA300 fix-hint
-->

## Summary
<!-- 1-3 bullet points describing what this PR changes and why. -->

-

## Linked issues
<!-- "Closes #123" / "Refs #456" -->

## Test plan
<!-- How was this verified? Required for any code change. -->

- [ ] `pytest tests/` passes locally
- [ ] `ruff check src tests` is clean
- [ ] If a new rule was added: `python tools/generate_rule_docs.py` was run and `docs/rules/` is committed
- [ ] Awesome-VBA baselines unchanged (or intentionally updated with rationale)

## New rules added
<!-- For each rule_id introduced, confirm: -->

- [ ] Registered in `src/rules.py`
- [ ] Doc page generated under `docs/rules/<id>.md`
- [ ] Fixture(s) added under `tests/samples/compile_errors/<category>/`
- [ ] Direct test in `tests/test_phaseN_*.py`

## Compatibility & migration
<!-- Any breaking changes? Default behaviour changes? Deprecations? -->
