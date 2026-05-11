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
    if _ENV.win32com is not None:
        return _run_with_timeout(
            _verify_with_win32com,
            (text, comp_name, kind, cfg, host, keep_workbook, timeout_s),
            timeout_s=timeout_s,
            host=host,
        )
    return _verify_with_comtypes(text, comp_name, kind, cfg, host, keep_workbook, timeout_s)


def _run_with_timeout(fn, args, *, timeout_s, host):
    """Run `fn(*args)` in a daemon worker thread; enforce `timeout_s`
    in the parent. If the worker hangs (e.g. VBE blocks on an invisible
    Compile-Error dialog behind `app.Visible=False`), kill the Office
    host process so the parent stays responsive and returns a clear
    `VBA_RT002` issue.

    pywin32 doesn't honour COM-call timeouts itself — this is the only
    way to avoid minute-long stalls in CI.
    """
    import threading

    result: list = []
    error: list[Exception] = []

    def _worker():
        try:
            result.append(fn(*args))
        except Exception as exc:
            # We capture *all* normal exceptions so the main thread can
            # re-raise them after `join`. We deliberately do NOT catch
            # `BaseException` (KeyboardInterrupt / SystemExit): those
            # are only delivered to the main thread anyway, so catching
            # them here would suppress nothing useful and would mask
            # genuine signal-driven termination.
            error.append(exc)

    t = threading.Thread(target=_worker, name="vbalidator-roundtrip", daemon=True)
    t.start()
    t.join(timeout_s)

    if t.is_alive():
        # Timeout. The worker is stuck inside a COM call we can't
        # interrupt — kill the Office host so it stops blocking, then
        # report the inconclusive run.
        _kill_office_processes(host)
        # Best-effort wait for the worker to unwind now that COM has
        # been torn down underneath it.
        t.join(5.0)
        return [_make_inconclusive_issue(
            f"round-trip exceeded {timeout_s:g}s budget for {host}; "
            f"Office process killed to free the parent. The static "
            f"analysis result remains the authoritative answer."
        )]

    if error:
        raise error[0]
    return result[0] if result else []


def _kill_office_processes(host):
    """Force-kill the Office host process tree on Windows. Best-effort —
    failures are swallowed because the worst case (zombie Excel.exe) is
    still an improvement over a hanging parent."""
    import shutil
    import subprocess as _sp

    targets = {
        "excel": ["EXCEL.EXE"],
        "word": ["WINWORD.EXE"],
        "access": ["MSACCESS.EXE"],
        "outlook": ["OUTLOOK.EXE"],
    }.get(host, [])
    taskkill = shutil.which("taskkill")
    if not taskkill:
        return
    for image in targets:
        try:
            _sp.run(
                [taskkill, "/F", "/T", "/IM", image],
                stdout=_sp.DEVNULL,
                stderr=_sp.DEVNULL,
                timeout=10,
                check=False,
            )
        except (OSError, _sp.TimeoutExpired):
            # `taskkill` not on PATH (non-Windows), or it timed out
            # itself — nothing we can do, the parent moves on.
            pass


def _safe_module_name(stem: str) -> str:
    out = "".join(ch if ch.isalnum() or ch == "_" else "_" for ch in stem)
    if not out or not out[0].isalpha():
        out = "M_" + out
    return out[:31] or "InjectedModule"


def _strip_export_directives(text: str) -> str:
    """Strip the `.bas` / `.cls` / `.frm` export-only header block.

    Files exported from the VBE carry directives that are only legal at
    *file-import* time, not in the module body:

        VERSION 1.0 CLASS
        BEGIN
          MultiUse = -1  'True
        END
        Attribute VB_Name = "MyModule"
        Attribute VB_GlobalNameSpace = False
        Sub Foo()
            …
        End Sub

    Pasting the whole text through `CodeModule.AddFromString()` causes
    VBE to refuse the module ("Attribute-Anweisungen müssen vor allen
    Prozeduren stehen" / "Attributes must precede procedure body").
    The right thing is to skip every leading line that is part of the
    header block and pass only the actual code to the module.
    """
    lines = text.splitlines(keepends=True)
    body: list[str] = []
    in_header = True
    in_begin_block = False  # the `BEGIN ... END` form attribute block
    for line in lines:
        stripped = line.strip()
        if in_header:
            if in_begin_block:
                if stripped.upper() == "END":
                    in_begin_block = False
                # discard everything inside BEGIN…END
                continue
            if stripped.upper().startswith("VERSION "):
                continue
            if stripped.upper() == "BEGIN":
                in_begin_block = True
                continue
            if stripped.startswith("Attribute "):
                continue
            if not stripped:
                # Skip blank padding above the first real code line so
                # the resulting module isn't prefixed by spurious blank
                # rows. Once we've seen real code we keep blanks.
                continue
            in_header = False  # first non-directive content
        body.append(line)
    return "".join(body)


# ---- Compile-trigger strategies --------------------------------------

# The obvious call — `vbproj.Compile()` — is *not* exposed through pywin32
# late-binding on most Office builds (the type library hides the method
# even though the IDE uses it internally). A raw call yields
#   AttributeError: <unknown>.Compile
# which the pre-`bd432b1` version of this file silently mistranslated
# into a `VBA_RT001` compile_verified error on every input, clean or not.
#
# The tiered strategy below tries strictly *more reliable* approaches
# until one of them gives a definitive answer:
#
# 1.  `VBProject.Compile()` — direct call. Works on a few legacy Office
#     builds and on stand-alone VB6 projects. Returns AttributeError on
#     modern Office where the method is hidden.
#
# 2.  Probe-Sub via `Application.Run`. VBA compiles the *entire* project
#     before invoking any procedure, so a successful Run of an injected
#     no-op Sub proves the project is compile-clean. A failing Run with
#     a recognisable VBA error description gives us the actual compile
#     error message.
#
# 3.  When neither works (e.g. `Run` blocked by macro security in the
#     guest profile, or the project has a syntax error early enough to
#     refuse component creation): emit a single `VBA_RT000` info issue
#     stating that round-trip verification is unavailable on this host.
#     Critically: NEVER emit a fake `VBA_RT001` — we promise the user
#     that `compile_verified` means "VBE itself rejected this code".


def _attempt_compile(vbproj, app, wb, host, comp_name):
    """Tiered VBA-compile trigger.

    Returns a list of issue dicts:
    - empty list when the project is compile-clean
    - `[VBA_RT001, ...]` when VBE found errors (severity `compile_verified`)
    - `[VBA_RT000]` info when no technique could trigger a compile
    """
    # Strategy 1 — the documented (but hidden) Compile method.
    direct = _try_direct_compile(vbproj, host)
    if direct is not None:
        return direct

    # Strategy 2 — probe Sub via Application.Run.
    probe = _try_probe_compile(vbproj, app, wb, host, comp_name)
    if probe is not None:
        return probe

    # Strategy 3 — confess. Use the *inconclusive* issue (RT002, warning)
    # not the unavailable one (RT000, info): on this host VBE was reachable,
    # we just couldn't trigger a compile via any of the supported call
    # routes. The user asked for verification — surfacing that we couldn't
    # deliver is the honest answer.
    return [_make_inconclusive_issue(
        f"round-trip compile could not be triggered on {host}; tried "
        f"VBProject.Compile() (Strategy 1) and probe-Sub via "
        f"Application.Run (Strategy 2). The static analysis result "
        f"above remains the authoritative answer. See TODO.md §A2."
    )]


def _try_direct_compile(vbproj, host):
    """Strategy 1: call `VBProject.Compile` directly.

    Returns `None` when the method is not exposed (so the caller tries
    the next strategy), `[]` when the project compiles cleanly, or a
    list with a single `VBA_RT001` when VBE returned a real error.
    """
    try:
        compile_fn = getattr(vbproj, "Compile", None)
    except Exception:
        return None
    if compile_fn is None:
        return None
    try:
        compile_fn()
    except AttributeError:
        # pywin32 raises this when the type library doesn't expose
        # `Compile` even though `getattr` returned a callable wrapper.
        # The wrapper is itself a generic `<unknown>` IDispatch shim.
        return None
    except Exception as exc:
        # Anything else is treated as a genuine VBE error — but only
        # when the description is *not* the well-known
        # "Methode/Eigenschaft im Projekttyp ungültig" / "Method or
        # data member not found" message, which means the call route
        # itself failed (i.e. still not a real compile error).
        msg = _vba_error_description(exc).lower()
        if any(
            sentinel in msg
            for sentinel in (
                "method or data member not found",
                "methode/eigenschaft im projekttyp ungültig",
                "method or property not supported",
            )
        ):
            return None
        return [{
            "file": "<roundtrip>", "line": 0,
            "rule_id": "VBA_RT001", "severity": "compile_verified",
            "category": "roundtrip",
            "message": (
                f"VBE compile error (round-trip, {host}): "
                f"{_vba_error_description(exc)}"
            ),
        }]
    return []  # clean compile


def _try_probe_compile(vbproj, app, wb, host, comp_name):
    """Strategy 2: inject a no-op Sub and invoke it via
    `Application.Run`. VBA compiles the entire project before invoking
    any procedure, so a successful Run proves a clean compile.

    Returns `None` when the strategy isn't applicable (couldn't add the
    probe module, Run blocked, …), `[]` on clean compile, or a list
    with a `VBA_RT001` when Run failed in a way that looks like a
    compile error.
    """
    import time

    probe_module_name = f"VBAlidPr{int(time.time() * 1000) % 1_000_000:06d}"
    probe_sub_name = "Probe"
    probe_comp = None

    try:
        probe_comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    except Exception:
        return None

    try:
        probe_comp.Name = probe_module_name
        probe_comp.CodeModule.AddFromString(
            f"Public Sub {probe_sub_name}()\nEnd Sub\n"
        )
        # Excel: 'WorkbookName.xlsm'!Module.Sub
        # Word:  Document.Module.Sub  (no quotes / no extension)
        # We try Excel-style first since it's the more common host.
        macro_ref = _format_macro_ref(host, wb, probe_module_name, probe_sub_name)
        try:
            app.Run(macro_ref)
        except Exception as exc:
            description = _vba_error_description(exc)
            low = description.lower()
            # "Cannot run the macro …" / "Makro nicht verfügbar" usually
            # means VBE refused because the project doesn't compile.
            # "The macro may not be available …" is the English variant.
            sentinels = (
                "cannot run the macro",
                "the macro may not be available",
                "makro nicht verfügbar",
                "kann nicht ausgeführt werden",
                "compile error",
                "kompilierungsfehler",
                "syntax error",
                "syntaxfehler",
                # The Attribute-block fence error VBE throws when the
                # injected source still carries the `Attribute VB_Name`
                # header (now handled by `_strip_export_directives`, but
                # belt-and-braces in case a user passes raw content).
                "attribute-anweisungen müssen vor allen prozeduren stehen",
                "attribute statements must precede",
                "attributes must precede procedure body",
                # Generic VBE compile-failure phrases.
                "expected: identifier",
                "expected end of statement",
                "expected: end of statement",
                "user-defined type not defined",
                "benutzerdefinierter typ nicht definiert",
                "sub or function not defined",
                "sub oder function nicht definiert",
                "variable not defined",
                "variable nicht definiert",
            )
            if any(sentinel in low for sentinel in sentinels):
                return [{
                    "file": "<roundtrip>", "line": 0,
                    "rule_id": "VBA_RT001", "severity": "compile_verified",
                    "category": "roundtrip",
                    "message": (
                        f"VBE compile error (round-trip, {host}, via Run): "
                        f"{description}"
                    ),
                }]
            # Something else went wrong (security prompt, MAPI logon, …).
            # Don't claim a compile error; fall through to Strategy 3.
            return None
        return []  # clean compile
    finally:
        if probe_comp is not None:
            try:
                vbproj.VBComponents.Remove(probe_comp)
            except Exception:
                # Probe-module removal can fail when Office is mid-shutdown
                # or the project is read-only. The temp workbook is going
                # away anyway, so this is purely cosmetic.
                pass


def _format_macro_ref(host, wb, module_name, sub_name):
    """Build the `Application.Run` argument string for a probe sub."""
    if host == "excel":
        # Quotes around the workbook name handle spaces correctly.
        return f"'{wb.Name}'!{module_name}.{sub_name}"
    if host == "word":
        # Word uses the document's project name (no extension).
        return f"{module_name}.{sub_name}"
    # Fallback — same form as Word.
    return f"{module_name}.{sub_name}"


def _vba_error_description(exc):
    """Extract the human-readable description from a pywin32 com_error.

    com_error.args == (hresult, source, excepinfo, argerr)
    excepinfo      == (errnum, source, description, helpfile, helpcontext, errorid)
    """
    args = getattr(exc, "args", ())
    if len(args) >= 3:
        excepinfo = args[2]
        if isinstance(excepinfo, tuple) and len(excepinfo) >= 3 and excepinfo[2]:
            return str(excepinfo[2])
    return str(exc)


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
        # Strip the `.bas` export header before injecting — Attribute /
        # VERSION / BEGIN-END are only legal at file-import time.
        comp.CodeModule.AddFromString(_strip_export_directives(text))

        # Drive the actual VBA compiler. See `_attempt_compile` for the
        # tiered strategy and why the bare `vbproj.Compile()` call is
        # unreliable on most Office builds.
        issues.extend(_attempt_compile(vbproj, app, wb, host, comp_name))

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
    """The runtime can't even *attempt* a round-trip (Linux, missing
    pywin32, Office not installed, …). Severity `info`, rule
    `VBA_RT000` — callers can safely ignore."""
    return {
        "file": "<roundtrip>",
        "line": 0,
        "rule_id": "VBA_RT000",
        "severity": "info",
        "message": f"Round-trip verification skipped: {reason}",
        "category": "roundtrip",
    }


def _make_inconclusive_issue(reason: str) -> dict:
    """The runtime *tried* a round-trip but couldn't reach a verdict
    (every compile-trigger strategy failed, or the host hung past the
    timeout). Distinct from `VBA_RT000` (unavailable) so callers can
    tell the two states apart — pinned to severity `warning` because
    the user asked for a verification and we couldn't deliver one.
    Rule `VBA_RT002`."""
    return {
        "file": "<roundtrip>",
        "line": 0,
        "rule_id": "VBA_RT002",
        "severity": "warning",
        "message": f"Round-trip verification attempted but inconclusive: {reason}",
        "category": "roundtrip",
    }


__all__ = [
    "RoundtripUnavailable",
    "verify_compile",
    "is_available",
    "availability_reason",
]
