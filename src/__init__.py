"""VBAlidator — Premium VBA static analysis & compile-safety prechecker.

Public API
----------
>>> from src import precheck
>>> result = precheck("MyModule.bas", host="excel")
>>> result.compile_safe, result.score
"""
from .api import PrecheckResult, precheck, precheck_source

__all__ = ["precheck", "precheck_source", "PrecheckResult"]
