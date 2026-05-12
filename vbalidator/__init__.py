"""Compatibility shim: lets `from vbalidator import precheck` work
after `pip install vbalidator`.

The actual implementation lives in the `src` package which still
ships in the wheel; this module re-exports the public surface so
the import name matches the PyPI distribution name (and matches
what the README, UAT script and AI-integration guide tell users
to type).

The `src/`-as-top-level package name is a known wart of the current
layout. A follow-up PR will move `src/` → `vbalidator/` proper and
delete this shim. Until then:

>>> from vbalidator import precheck
>>> result = precheck("Module1.bas", host="excel")
>>> result.compile_safe, result.score

is equivalent to `from src import precheck`.

`vbalidator.<submodule>` works too (api, scoring, reporting, rules,
roundtrip) via the sys.modules aliases below — important for tooling
that imports inner modules, e.g.
`from vbalidator.rules import all_rules`.
"""
from __future__ import annotations

import sys

# Pull every submodule we want to alias eagerly so importers like
# `from vbalidator.api import precheck` resolve without surprises.
from src import (  # noqa: F401  (re-exported public surface)
    PrecheckResult,
    __version__,
    precheck,
    precheck_source,
)
from src import (
    api as _api,
    analyzer as _analyzer,
    config as _config,
    lexer as _lexer,
    parser as _parser,
    preprocessor as _preprocessor,
    reporting as _reporting,
    roundtrip as _roundtrip,
    rules as _rules,
    scoring as _scoring,
)

# Make `from vbalidator.<sub> import …` resolve to the same object
# `from src.<sub> import …` would. `setdefault` so a real
# `vbalidator/<sub>.py` (if added in the future migration) wins.
for _name, _mod in {
    "api": _api,
    "analyzer": _analyzer,
    "config": _config,
    "lexer": _lexer,
    "parser": _parser,
    "preprocessor": _preprocessor,
    "reporting": _reporting,
    "roundtrip": _roundtrip,
    "rules": _rules,
    "scoring": _scoring,
}.items():
    sys.modules.setdefault(f"vbalidator.{_name}", _mod)

del _api, _analyzer, _config, _lexer, _parser, _preprocessor
del _reporting, _roundtrip, _rules, _scoring
del _name, _mod, sys

__all__ = ["precheck", "precheck_source", "PrecheckResult", "__version__"]
