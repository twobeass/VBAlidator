"""Public Python API for VBAlidator — `precheck()` returns a structured
result that is directly usable as a post-processing step behind LLM-VBA
generators (LangChain / Anthropic SDK / OpenAI Agents).

Typical usage
-------------

>>> from vbalidator import precheck
>>> result = precheck("Module1.bas", host="excel")
>>> if not result.compile_safe:
...     for err in result.errors:
...         print(err["message"])
>>> result.score        # 0..100
>>> result.json()       # canonical JSON v2 report

Inputs may be a string of VBA source, a single file path, or a directory.
The output `PrecheckResult` is a thin dataclass over the list of
normalised issues so callers can both filter (.errors / .warnings / .info)
and serialise (.json()).
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .analyzer import Analyzer
from .config import Config
from .lexer import Lexer
from .parser import VBAParser, FormParser
from .preprocessor import Preprocessor
from .reporting import build_report_v2, normalize_issues
from .scoring import compute_score, is_compile_safe


_VBA_EXTS = (".bas", ".cls", ".frm")


@dataclass
class PrecheckResult:
    """Structured result of a precheck run."""

    score: int = 100
    compile_safe: bool = True
    issues: list[dict] = field(default_factory=list)  # normalised
    files_scanned: int = 0
    score_breakdown: dict = field(default_factory=dict)

    @property
    def errors(self) -> list[dict]:
        return [i for i in self.issues if i.get("severity") == "error"]

    @property
    def warnings(self) -> list[dict]:
        return [i for i in self.issues if i.get("severity") == "warning"]

    @property
    def info(self) -> list[dict]:
        return [i for i in self.issues if i.get("severity") == "info"]

    def json(self) -> dict:
        """Return the canonical JSON v2 report."""
        return build_report_v2(
            issues=self.issues,
            files_scanned=self.files_scanned,
            score=self.score,
            compile_safe=self.compile_safe,
            score_breakdown=self.score_breakdown,
        )

    def __bool__(self) -> bool:
        # Truthy when the source is compile-safe.
        return self.compile_safe


def _iter_input_files(source: str | os.PathLike, inline_name: str = "<inline>") -> tuple[list[tuple[str, str]], int]:
    """Resolve `source` to a list of (filename, content) pairs and a
    stable display root (used for relative paths)."""
    if isinstance(source, (str, os.PathLike)) and (
        isinstance(source, os.PathLike) or os.sep in str(source) or len(str(source)) < 4096
    ):
        # Heuristic: treat short non-newline strings as path-like.
        s = str(source)
        if "\n" not in s and (os.path.isfile(s) or os.path.isdir(s)):
            p = Path(s)
            if p.is_dir():
                files = []
                for root, _, fnames in os.walk(p):
                    for f in fnames:
                        if f.lower().endswith(_VBA_EXTS):
                            full = os.path.join(root, f)
                            with open(full, "r", encoding="latin-1") as fh:
                                files.append((os.path.relpath(full, p), fh.read()))
                return files, len(files)
            else:
                with open(p, "r", encoding="latin-1") as fh:
                    return [(p.name, fh.read())], 1
    # Inline source string.
    return [(inline_name, str(source))], 1


def _is_path_like(source) -> bool:
    """Cheap pre-check that mirrors `_iter_input_files`'s heuristic
    without re-reading the file. Returns True when `source` looks like a
    filesystem path that exists; False for inline source strings.
    """
    if isinstance(source, os.PathLike):
        return True
    if not isinstance(source, str):
        return False
    if "\n" in source:
        return False
    try:
        return os.path.isfile(source) or os.path.isdir(source)
    except (TypeError, ValueError, OSError):
        return False


def _autodetect_vba_model(source) -> Path | None:
    """Look for a `vba_model.json` next to the input and in the CWD.

    Search order (first hit wins):
    1. `<input_dir>/vba_model.json` when `source` is a directory.
    2. `<input_file_dir>/vba_model.json` when `source` is a file.
    3. `./vba_model.json` (current working directory).
    """
    candidates: list[Path] = []
    src_path = Path(str(source))
    if src_path.is_dir():
        candidates.append(src_path / "vba_model.json")
    elif src_path.is_file():
        candidates.append(src_path.parent / "vba_model.json")
    candidates.append(Path.cwd() / "vba_model.json")
    for c in candidates:
        if c.is_file():
            return c
    return None


# Auto-layer table: which companion model to layer when the scan set
# references it. Each entry is (model filename, regex to look for in
# any source file, restrict-to-extension or None for "any file").
#
# Pattern design note: we trigger on the *namespace* name (`MSForms`,
# `Scripting`, …) rather than the bare class names (`Dictionary`,
# `UserForm`, …) because the unqualified class names are too generic —
# they clash with project-internal identifiers in real libraries.
import re as _re_aux

_AUTO_LAYER_RULES: list[tuple[str, "_re_aux.Pattern[str]", str | None]] = [
    ("mscomctl.json", _re_aux.compile(r"\b(?:MS)?Comctl(?:Lib)?\b", _re_aux.IGNORECASE), ".frm"),
    ("msforms.json", _re_aux.compile(r"\bMSForms\b", _re_aux.IGNORECASE), None),
    # ProgID-style auto-layers — match the namespace prefix of CreateObject
    # strings or `As Scripting.X` declarations. Spelled out verbatim because
    # VBA is case-insensitive but real-world capitalisation drifts.
    ("scripting.json", _re_aux.compile(r"\bScripting\.(?:Dictionary|FileSystemObject)\b", _re_aux.IGNORECASE), None),
    ("vbscript_regexp.json", _re_aux.compile(r"\bVBScript\.Reg[Ee]xp\b", _re_aux.IGNORECASE), None),
    ("wscript_shell.json", _re_aux.compile(r"\bWScript\.Shell\b", _re_aux.IGNORECASE), None),
    ("shell_application.json", _re_aux.compile(r"\bShell\.Application\b", _re_aux.IGNORECASE), None),
]


def apply_auto_layers(config: Config, files: list[tuple[str, str]]) -> list[str]:
    """Layer companion models on top of the standard / host model when
    the scan set references them. Returns the list of layered model
    filenames (mostly for logging / tests). Shared between `precheck()`
    and the test conftest pipeline so the two stay in sync."""
    models_dir = Path(__file__).resolve().parent / "models"
    layered: list[str] = []
    for model_name, pat, ext_filter in _AUTO_LAYER_RULES:
        candidates = files
        if ext_filter is not None:
            candidates = [
                (fn, ct) for fn, ct in files
                if os.path.splitext(fn)[1].lower() == ext_filter
            ]
        if any(pat.search(content) for _, content in candidates):
            path = models_dir / model_name
            if path.is_file():
                config.load_model(str(path))
                layered.append(model_name)
    return layered


def _load_host_model(config: Config, host: str | None) -> bool:
    """Load `models/<host>.json` if it exists. Return True if loaded.
    Silent no-op when host is None or the file does not exist (the user
    might supply a custom model via `--model`)."""
    if not host:
        return False
    base = Path(__file__).resolve().parent / "models"
    candidate = base / f"{host.lower()}.json"
    if not candidate.is_file():
        return False
    config.load_model(str(candidate))
    return True


def precheck(
    source: str | os.PathLike,
    *,
    host: str | None = None,
    model_path: str | os.PathLike | None = None,
    defines: dict[str, Any] | None = None,
    strict: bool = True,
    module_type: str | None = None,
    roundtrip: bool = False,
) -> PrecheckResult:
    """Run the full VBAlidator pipeline.

    Parameters
    ----------
    source
        Either an inline VBA source string, a path to a `.bas/.cls/.frm`
        file, or a directory walked recursively.
    host
        One of `excel`, `word`, `access`, `outlook`, `vba_runtime`. When
        provided the matching `models/<host>.json` is auto-loaded so the
        user does not need to run the model exporter first.
    model_path
        Path to a custom JSON object model. Layered on top of the host
        model and the bundled `std_model.json`.
    defines
        Conditional-compilation constants (e.g. ``{"WIN64": False}``).
    strict
        When False, severity=='warning' findings do not affect
        `compile_safe`. Errors always do.
    module_type
        Override module type when `source` is an inline string. Defaults
        to "Module" / "Class" / "Form" inferred from the extension.
    """
    config = Config()
    if defines:
        for k, v in defines.items():
            config.definitions[k.upper()] = v
    if host:
        _load_host_model(config, host)
    if model_path:
        config.load_model(str(model_path))
    elif _is_path_like(source):
        # Auto-load `vba_model.json` if present next to the input or in
        # the current working directory. Keeps the documented "drop a
        # vba_model.json next to your code" UX from Phase 0–3.
        auto = _autodetect_vba_model(source)
        if auto is not None:
            config.load_model(str(auto))

    files, n_files = _iter_input_files(source)

    apply_auto_layers(config, files)

    analyzer = Analyzer(config)

    for filename, content in files:
        ext = os.path.splitext(filename)[1].lower()
        if module_type is not None and len(files) == 1:
            mtype = module_type
        elif ext == ".cls":
            mtype = "Class"
        elif ext == ".frm":
            mtype = "Form"
        else:
            mtype = "Module"

        # Form: scrape implicit controls before lexing the code part.
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
            analyzer.errors.append(lex_err.to_dict(filename=filename))

        pp = Preprocessor(tokens, config.definitions)
        processed_tokens = list(pp.process())

        parser = VBAParser(processed_tokens, filename=filename)
        module_node = parser.parse_module()
        module_node.filename = filename
        module_node.module_type = mtype
        for syn in parser.errors:
            analyzer.errors.append(syn)
        if ext == ".frm":
            module_node.variables.extend(controls)
        analyzer.add_module(module_node)

    raw_issues = analyzer.analyze()

    # Phase 4.5 — optional dynamic verification through Office COM.
    if roundtrip:
        try:
            from .roundtrip import is_available, availability_reason, verify_compile
            if not is_available():
                raw_issues.append({
                    "file": "<roundtrip>", "line": 0,
                    "rule_id": "VBA_RT000", "severity": "info",
                    "category": "roundtrip",
                    "message": f"Round-trip verification unavailable: {availability_reason()}",
                })
            else:
                # Round-trip every input file individually so the per-file
                # error attribution stays correct.
                for filename, content in files:
                    rt_issues = verify_compile(
                        content,
                        host=(host or "excel"),
                    )
                    # Re-attribute the file name (verify_compile uses the
                    # injected component name by default).
                    for i in rt_issues:
                        i["file"] = filename
                    raw_issues.extend(rt_issues)
        except Exception as exc:  # pragma: no cover — defence-in-depth
            raw_issues.append({
                "file": "<roundtrip>", "line": 0,
                "rule_id": "VBA_RT000", "severity": "info",
                "category": "roundtrip",
                "message": f"Round-trip verification crashed: {exc}",
            })

    issues = normalize_issues(raw_issues)

    if not strict:
        # Drop warnings + info from the gating set; keep them in the
        # report so the caller can still see them.
        gating = [i for i in issues if i.get("severity") == "error"]
    else:
        gating = issues

    score, breakdown = compute_score(gating)
    safe = is_compile_safe(gating)

    return PrecheckResult(
        score=score,
        compile_safe=safe,
        issues=issues,
        files_scanned=n_files,
        score_breakdown=breakdown,
    )


def precheck_source(
    source: str,
    *,
    name: str = "<inline>",
    host: str | None = None,
    model_path: str | os.PathLike | None = None,
    defines: dict[str, Any] | None = None,
    module_type: str | None = None,
) -> PrecheckResult:
    """Convenience wrapper for an inline source string with a custom
    display name (otherwise `precheck` emits `<inline>` in the report).
    """
    config = Config()
    if defines:
        for k, v in defines.items():
            config.definitions[k.upper()] = v
    if host:
        _load_host_model(config, host)
    if model_path:
        config.load_model(str(model_path))

    analyzer = Analyzer(config)
    lexer = Lexer(source)
    tokens = list(lexer.tokenize())
    for lex_err in lexer.errors:
        analyzer.errors.append(lex_err.to_dict(filename=name))

    pp = Preprocessor(tokens, config.definitions)
    processed = list(pp.process())
    parser = VBAParser(processed, filename=name)
    module_node = parser.parse_module()
    module_node.filename = name
    module_node.module_type = module_type or "Module"
    for syn in parser.errors:
        analyzer.errors.append(syn)
    analyzer.add_module(module_node)

    issues = normalize_issues(analyzer.analyze())
    score, breakdown = compute_score(issues)
    return PrecheckResult(
        score=score,
        compile_safe=is_compile_safe(issues),
        issues=issues,
        files_scanned=1,
        score_breakdown=breakdown,
    )


__all__ = ["precheck", "precheck_source", "PrecheckResult"]
