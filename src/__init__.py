"""VBAlidator — Premium VBA static analysis & compile-safety prechecker.

Public API
----------
>>> from src import precheck
>>> result = precheck("MyModule.bas", host="excel")
>>> result.compile_safe, result.score
"""
__version__ = "1.1.1"

from .api import PrecheckResult, precheck, precheck_source

__all__ = ["precheck", "precheck_source", "PrecheckResult", "__version__"]
