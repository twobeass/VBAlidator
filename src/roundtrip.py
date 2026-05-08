"""Round-Trip Verification — drive the actual VBE compiler via COM and
diff its result against the static analyser. Phase 4.5 of the roadmap.

Why
---
VBAlidator is a static analyser; it ships with deterministic rules that
match VBE's compile-time behaviour as closely as possible. Round-trip
verification provides a *dynamic* second opinion: we inject the source
into a temporary `.xlsm`/`.docm`, drive Office headlessly through COM,
and call `VBProject.Compile`. Any compile errors VBE reports are then
reflected back as `severity='compile_verified'` issues so callers can
distinguish "the static analyser thinks this is broken" from "the actual
VBA compiler refuses this".

Platform requirements
---------------------
- Windows
- Microsoft Office (Excel for `.xlsm`, Word for `.docm`) installed and
  COM-registered
- The "Trust access to the VBA project object model" setting enabled in
  Office Trust Center (otherwise `VBProject` is not exposed)
- pywin32 OR comtypes installed (this module imports lazily)

On non-Windows hosts the module degrades gracefully: callers that try to
use it get a clear `RoundtripUnavailable` exception with the reason; the
CLI emits a single warning and continues without dynamic verification.

Public surface
--------------
- `is_available()` — True when the platform can run a round-trip.
- `verify_compile(source, host="excel", timeout_s=30) → list[Issue]`
  Returns a list of issues with severity `'compile_verified'` and
  rule_id `VBA_RT001`. The list is empty when VBE compiles cleanly.

The function is intentionally side-effect-free outside of the temporary
file: Office is launched with `Visible=False`, `DisplayAlerts=False`,
the temp file is deleted, and the Office app is quit even on errors.
"""
from __future__ import annotations

import os
import sys
import tempfile
from dataclasses import dataclass


class RoundtripUnavailable(RuntimeError):
    """Raised when round-trip verification cannot run on this host."""


@dataclass
class _ComEnv:
    """Holds the optional COM entry-point — populated lazily."""
    win32com: object = None
    comtypes: object = None


_ENV = _ComEnv()


def _platform_supported() -> tuple[bool, str]:
    if sys.platform != "win32":
        return False, f"round-trip requires Windows; running on {sys.platform!r}"
    return True, ""


def _import_com() -> str:
    """Try to import a COM dispatcher. Mutates _ENV. Returns the engine
    name or '' if neither library is installed."""
    if _ENV.win32com is not None or _ENV.comtypes is not None:
        return "win32com" if _ENV.win32com else "comtypes"
    try:
        import win32com.client as wc  # type: ignore
        _ENV.win32com = wc
        return "win32com"
    except ImportError:
        pass
    try:
        import comtypes.client as ct  # type: ignore
        _ENV.comtypes = ct
        return "comtypes"
    except ImportError:
        pass
    return ""


def is_available() -> bool:
    """True when the current host can run a round-trip verification.

    This does *not* guarantee Office is installed — only that the
    platform / Python bindings are present. The actual COM call may
    still fail with a clearer error.
    """
    ok, _ = _platform_supported()
    if not ok:
        return False
    return bool(_import_com())


def availability_reason() -> str:
    """Return a human-readable explanation when `is_available()` is False."""
    ok, why = _platform_supported()
    if not ok:
        return why
    if not _import_com():
        return (
            "neither pywin32 nor comtypes is installed; "
            "`pip install pywin32` to enable round-trip verification"
        )
    return ""


# ---- Source injection helpers ----------------------------------------

_HOST_CONFIG = {
    "excel": {
        "progid": "Excel.Application",
        "ext": ".xlsm",
        "add_method": "Add",
        "save_format": 52,  # xlOpenXMLWorkbookMacroEnabled
        "vbproject_path": ("ActiveWorkbook", "VBProject"),
        "module_kind": 1,   # vbext_ct_StdModule
    },
    "word": {
        "progid": "Word.Application",
        "ext": ".docm",
        "add_method": "Add",
        "save_format": 13,  # wdFormatXMLDocumentMacroEnabled
        "vbproject_path": ("ActiveDocument", "VBProject"),
        "module_kind": 1,
    },
}


def _module_kind_for_extension(ext: str) -> int:
    """Map a file extension to the vbext_ComponentType used by VBE."""
    e = ext.lower()
    if e == ".cls":
        return 2  # vbext_ct_ClassModule
    if e == ".frm":
        return 3  # vbext_ct_MSForm — not supported via inject; falls back to StdModule
    return 1      # vbext_ct_StdModule


def verify_compile(
    source: str | os.PathLike,
    *,
    host: str = "excel",
    timeout_s: float = 30,
    keep_workbook: bool = False,
) -> list[dict]:
    """Run a real VBE compile pass on `source`.

    Parameters
    ----------
    source : str or path-like
        Either a single VBA file path or an inline source string.
    host : "excel" | "word"
        Which Office app to drive. Excel is the most universal target.
    timeout_s : float
        Soft upper bound for COM calls (informational; pywin32 itself
        does not honour this directly, but it is included in the
        emitted error metadata for diagnostics).
    keep_workbook : bool
        Leave the temporary host file on disk — useful for debugging.

    Returns
    -------
    list of issue dicts (`severity='compile_verified'`, `rule_id='VBA_RT001'`).
    Empty list when VBE compiles cleanly.
    """
    if not is_available():
        raise RoundtripUnavailable(availability_reason())

    if host not in _HOST_CONFIG:
        raise ValueError(f"Unsupported round-trip host: {host!r}. Choose 'excel' or 'word'.")

    cfg = _HOST_CONFIG[host]

    # Resolve source → (text, name).
    if os.path.exists(str(source)):
        path = str(source)
        with open(path, "r", encoding="latin-1") as f:
            text = f.read()
        comp_name = _safe_module_name(os.path.splitext(os.path.basename(path))[0])
        kind = _module_kind_for_extension(os.path.splitext(path)[1])
    else:
        text = str(source)
        comp_name = "InjectedModule"
        kind = cfg["module_kind"]

    # win32com is the more reliable backend. We only fall back to comtypes
    # when win32com isn't present.
    issues: list[dict] = []
    if _ENV.win32com is not None:
        issues = _verify_with_win32com(text, comp_name, kind, cfg, host, keep_workbook, timeout_s)
    else:
        issues = _verify_with_comtypes(text, comp_name, kind, cfg, host, keep_workbook, timeout_s)
    return issues


def _safe_module_name(stem: str) -> str:
    out = "".join(ch if ch.isalnum() or ch == "_" else "_" for ch in stem)
    if not out or not out[0].isalpha():
        out = "M_" + out
    return out[:31] or "InjectedModule"


def _verify_with_win32com(
    text: str,
    comp_name: str,
    kind: int,
    cfg: dict,
    host: str,
    keep_workbook: bool,
    timeout_s: float,
) -> list[dict]:
    wc = _ENV.win32com
    app = wc.Dispatch(cfg["progid"])
    app.Visible = False
    if hasattr(app, "DisplayAlerts"):
        try:
            app.DisplayAlerts = False
        except Exception:
            # Best-effort: some hosts don't expose DisplayAlerts on every
            # build / SKU, but the call still attempts the property. A
            # failure here doesn't affect compile verification.
            pass

    tmp = tempfile.NamedTemporaryFile(suffix=cfg["ext"], delete=False)
    tmp.close()
    tmp_path = tmp.name
    issues: list[dict] = []

    try:
        if host == "excel":
            wb = app.Workbooks.Add()
        elif host == "word":
            wb = app.Documents.Add()
        else:
            raise ValueError(host)

        try:
            wb.SaveAs(tmp_path, cfg["save_format"])
        except Exception as exc:
            issues.append(_make_unavailable_issue(
                f"could not save temporary {host} workbook: {exc}"
            ))
            return issues

        # Inject the source as a fresh component.
        try:
            vbproj = wb.VBProject
        except Exception as exc:
            issues.append(_make_unavailable_issue(
                f"VBProject access denied — enable 'Trust access to the VBA "
                f"project object model' in Office Trust Center. ({exc})"
            ))
            return issues

        comp = vbproj.VBComponents.Add(kind)
        comp.Name = comp_name
        comp.CodeModule.AddFromString(text)

        # Compile and capture errors.
        try:
            vbproj.Compile()  # raises if compile fails
        except Exception as exc:
            issues.append({
                "file": comp_name,
                "line": 0,
                "rule_id": "VBA_RT001",
                "severity": "compile_verified",
                "message": f"VBE compile error (round-trip, {host}): {exc}",
                "category": "roundtrip",
            })

    finally:
        # All three teardown calls are deliberately best-effort. The
        # finally block runs even when the prior block already failed
        # (Office crash, COM disconnect, file lock), so re-raising here
        # would mask the original error and is never useful.
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            # Workbook may have been closed by the host (`Quit` callback)
            # or never opened cleanly. Either way, nothing actionable.
            pass
        try:
            app.Quit()
        except Exception:
            # Office may already be shutting down; ignore.
            pass
        if not keep_workbook:
            try:
                os.unlink(tmp_path)
            except OSError:
                # File handle may still be held by Office on a slow
                # machine; the temp dir is cleaned up by the OS later.
                pass

    return issues


def _verify_with_comtypes(
    text: str,
    comp_name: str,
    kind: int,
    cfg: dict,
    host: str,
    keep_workbook: bool,
    timeout_s: float,
) -> list[dict]:
    # comtypes path is left as an explicit unavailability for now —
    # win32com is by far the more common deployment.
    return [_make_unavailable_issue(
        "comtypes round-trip backend not yet implemented; install pywin32"
    )]


def _make_unavailable_issue(reason: str) -> dict:
    return {
        "file": "<roundtrip>",
        "line": 0,
        "rule_id": "VBA_RT000",
        "severity": "info",
        "message": f"Round-trip verification skipped: {reason}",
        "category": "roundtrip",
    }


__all__ = [
    "RoundtripUnavailable",
    "verify_compile",
    "is_available",
    "availability_reason",
]
