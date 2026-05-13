"""Shared pytest fixtures and pipeline helpers for VBAlidator tests."""
from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

# Make the project root importable so `from src import ...` works in tests.
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest

from src.analyzer import Analyzer
from src.config import Config
from src.lexer import Lexer
from src.parser import FormParser, VBAParser
from src.preprocessor import Preprocessor


@dataclass
class AnalysisResult:
    errors: list[dict[str, Any]] = field(default_factory=list)
    lexer_errors: list[Any] = field(default_factory=list)

    @property
    def messages(self) -> list[str]:
        return [e.get("message", "") for e in self.errors]

    def has_message_containing(self, substring: str) -> bool:
        substring_low = substring.lower()
        return any(substring_low in m.lower() for m in self.messages)


def _module_type_for(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".cls":
        return "Class"
    if ext == ".frm":
        return "Form"
    return "Module"


def run_pipeline_on_files(
    paths: list[Path],
    extra_defines: dict | None = None,
    host: str | None = None,
) -> AnalysisResult:
    """Run the full Lexer→Preprocessor→Parser→Analyzer pipeline.

    Mirrors src/main.py but works on an explicit file list instead of a folder
    walk and returns a structured result for assertion in tests.

    `host` layers the bundled `src/models/<host>.json` on top of std_model —
    same path the CLI takes for `--host excel|word|access|outlook|visio`.
    Defaults to None (std_model only) for backward-compat with the existing
    tests that don't pass it.
    """
    config = Config()
    if extra_defines:
        for k, v in extra_defines.items():
            config.definitions[k.upper()] = v
    if host:
        model_path = ROOT / "src" / "models" / f"{host}.json"
        if model_path.is_file():
            config.load_model(str(model_path))

    analyzer = Analyzer(config)
    lexer_errors: list[Any] = []

    for path in paths:
        path = Path(path)
        with open(path, "r", encoding="latin-1") as f:
            content = f.read()

        ext = path.suffix.lower()
        module_type = _module_type_for(path)

        controls = []
        code_content = content
        if ext == ".frm":
            fp = FormParser()
            controls = fp.parse(content)
            import re as _re
            match = _re.search(r"Attribute\s+VB_Name", content)
            if match:
                code_content = content[match.start():]

        lexer = Lexer(code_content)
        tokens = list(lexer.tokenize())
        for lex_err in lexer.errors:
            analyzer.errors.append(lex_err.to_dict(filename=path.name))
            lexer_errors.append(lex_err)

        pp = Preprocessor(tokens, config.definitions)
        processed_tokens = list(pp.process())

        parser = VBAParser(processed_tokens, filename=path.name)
        module_node = parser.parse_module()
        module_node.filename = path.name
        module_node.module_type = module_type

        for syn_err in parser.errors:
            analyzer.errors.append(syn_err)

        if ext == ".frm":
            module_node.variables.extend(controls)

        analyzer.add_module(module_node)

    errors = analyzer.analyze()
    return AnalysisResult(errors=list(errors), lexer_errors=lexer_errors)


def run_pipeline_on_source(code: str, module_type: str = "Module") -> AnalysisResult:
    """Run pipeline on a single in-memory VBA source string."""
    config = Config()
    analyzer = Analyzer(config)

    lexer = Lexer(code)
    tokens = list(lexer.tokenize())
    for lex_err in lexer.errors:
        analyzer.errors.append(lex_err.to_dict(filename="<inline>"))

    pp = Preprocessor(tokens, config.definitions)
    processed_tokens = list(pp.process())

    parser = VBAParser(processed_tokens, filename="<inline>")
    module_node = parser.parse_module()
    module_node.filename = "<inline>"
    module_node.module_type = module_type
    for syn_err in parser.errors:
        analyzer.errors.append(syn_err)
    analyzer.add_module(module_node)

    return AnalysisResult(errors=list(analyzer.analyze()), lexer_errors=list(lexer.errors))


@pytest.fixture
def run_source():
    """Pytest fixture exposing run_pipeline_on_source as a callable."""
    return run_pipeline_on_source


@pytest.fixture
def run_files():
    """Pytest fixture exposing run_pipeline_on_files as a callable."""
    return run_pipeline_on_files


@pytest.fixture(scope="session")
def repo_root() -> Path:
    return ROOT


@pytest.fixture(scope="session")
def samples_dir(repo_root) -> Path:
    return repo_root / "tests" / "samples"


@pytest.fixture(scope="session")
def awesome_vba_dir(repo_root) -> Path:
    return repo_root / "tests" / "awesome_vba"
