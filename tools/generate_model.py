#!/usr/bin/env python3
"""Generate a VBAlidator object-model JSON from a host's COM type
libraries.

Workflow
--------

1. Run `VBA_Model_Exporter.bas` inside the VBE of your target host
   (Excel / Word / Access / PowerPoint / Outlook / Visio / AutoCAD).
   It writes a `vba_references.json` file listing every referenced
   library with its GUID + version.

2. Run this script:

      python tools/generate_model.py path/to/vba_references.json -o vba_model.json

   It uses `comtypes` to introspect every library: extracts dispinterfaces,
   coclasses, enums and module-level constants and emits a single
   `vba_model.json` consumable via `vbalidator --model vba_model.json`
   or `precheck(model_path=…)`.

3. The generated model is layered on top of the bundled `std_model.json`
   and any `--host` model, so it only needs to cover host-specific
   types — VBA built-ins are already there.

Platform: Windows + the target host installed (comtypes drives COM).
On non-Windows the script exits with an actionable error.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any

LOG = logging.getLogger("vbalidator.gen_model")


def _import_comtypes():
    try:
        import comtypes  # noqa: F401  (re-exported via comtypes.client)
        import comtypes.client as cc
    except ImportError as exc:  # pragma: no cover — Windows-only
        sys.stderr.write(
            "Error: 'comtypes' is required.\n"
            "  pip install comtypes\n"
            f"  (import error: {exc})\n"
        )
        sys.exit(1)
    return cc


# ---------------------------------------------------------------- helpers


def _load_module(cc, ref: dict[str, Any]):
    """Resolve a library entry into a comtypes module. First by GUID +
    version, then by absolute path, then fall through to None."""
    guid = ref.get("guid")
    major = ref.get("major")
    minor = ref.get("minor")
    path = ref.get("path", "")

    if guid:
        try:
            return cc.GetModule((guid, major or 1, minor or 0))
        except (OSError, ValueError, AttributeError) as exc:
            LOG.debug("GetModule(GUID) failed for %s: %s", ref.get("name"), exc)
    if path and os.path.exists(path):
        try:
            return cc.GetModule(path)
        except (OSError, ValueError, AttributeError) as exc:
            LOG.debug("GetModule(path) failed for %s: %s", ref.get("name"), exc)
    return None


def _extract_member_args(method_tuple: tuple) -> list[dict[str, Any]]:
    """comtypes serialises each method as
    `(['flags'], 'ReturnType', 'Name', (['in'], 'ArgType', 'ArgName'), …)`.
    Walk that and surface a uniform list of arg dicts."""
    args_info: list[dict[str, Any]] = []
    if len(method_tuple) <= 3:
        return args_info

    candidates = []
    for i in range(3, len(method_tuple)):
        item = method_tuple[i]
        # Some bindings pack args as a nested tuple-of-tuples.
        if isinstance(item, tuple) and item and isinstance(item[0], tuple):
            candidates.extend(item)
        else:
            candidates.append(item)

    for arg_def in candidates:
        if not (isinstance(arg_def, tuple) and len(arg_def) >= 3):
            continue
        flags = arg_def[0]
        if not isinstance(flags, (list, tuple)):
            continue

        # comtypes wraps types as classes (HRESULT, c_int, …) — use the
        # class name when present, otherwise the repr (good enough for
        # documentation purposes).
        type_obj = arg_def[1]
        arg_type = getattr(type_obj, "__name__", None) or str(type_obj)

        # `str()` is essentially infallible on Python objects; if a COM
        # binding ever exposed a `__str__` that raised we'd rather see
        # the traceback than swallow it silently.
        arg_name = str(arg_def[2])

        if "out" in flags or ("in" not in flags and "out" not in flags):
            mechanism = "ByRef"
        else:
            mechanism = "ByVal"

        args_info.append({
            "name": arg_name,
            "type": arg_type,
            "mechanism": mechanism,
            "is_optional": "opt" in flags,
        })
    return args_info


def _clean_member_name(name: str) -> str:
    """Strip comtypes' `_get_` / `_set_` / `_put_` decorators."""
    for prefix in ("_get_", "_set_", "_put_"):
        if name.startswith(prefix):
            return name[len(prefix):]
    return name


# ---------------------------------------------------------------- passes


def _process_interfaces_and_enums(model: dict, mod, ref_name: str) -> None:
    """Pass 1 — register every dispinterface as a class, every enum as an
    enum, and promote the VBA library's globals."""
    for attr_name in dir(mod):
        try:
            attr = getattr(mod, attr_name)
        except (AttributeError, OSError):
            continue

        # Dispinterface / Interface
        if isinstance(attr, type) and (
            hasattr(attr, "_methods_") or hasattr(attr, "_disp_methods_")
        ):
            type_name = str(attr_name)
            cls = model["classes"].setdefault(
                type_name, {"type": "Class", "members": {}},
            )
            all_methods = list(getattr(attr, "_methods_", [])) + list(
                getattr(attr, "_disp_methods_", [])
            )
            for m in all_methods:
                if len(m) < 3:
                    continue
                raw_name = (
                    m[2] if isinstance(m[2], str) else
                    (m[2][0] if isinstance(m[2], tuple) and m[2] and isinstance(m[2][0], str) else None)
                )
                if not raw_name and len(m) >= 2 and isinstance(m[1], str):
                    raw_name = m[1]
                if not raw_name:
                    continue

                clean = _clean_member_name(str(raw_name))
                args_info = _extract_member_args(m)
                member = {"type": "Variant"}
                if args_info:
                    member["args"] = args_info
                    member["min_args"] = sum(1 for a in args_info if not a["is_optional"])
                    member["max_args"] = len(args_info)
                cls["members"][clean] = member

                # The VBA stdlib library registers as ref_name == "VBA";
                # promote its members directly into the global scope.
                if ref_name == "VBA" and not clean.startswith("_"):
                    model["globals"].setdefault(clean, member)
            continue

        # Top-level integer constant (e.g. visNone)
        if isinstance(attr, int) and not attr_name.startswith("_"):
            model["globals"].setdefault(
                attr_name, {"type": "Long", "kind": "Constant"},
            )
            continue

        # Constants container — class with int attributes
        if isinstance(attr, type):
            enum_values: dict[str, int] = {}
            for e_name in dir(attr):
                if e_name.startswith("_"):
                    continue
                try:
                    val = getattr(attr, e_name)
                except (AttributeError, OSError):
                    continue
                if isinstance(val, int):
                    enum_values[e_name] = val
            if enum_values:
                enum_name = str(attr_name)
                bucket = model["enums"].setdefault(enum_name, {})
                bucket.update(enum_values)
                for e_k in enum_values:
                    model["globals"].setdefault(
                        e_k, {"type": "Long", "kind": "EnumItem"},
                    )


def _process_coclasses(model: dict, mod) -> None:
    """Pass 2 — copy every CoClass's default-interface members onto the
    CoClass entry (so `Set x As Workbook` resolves the same as the
    underlying `IWorkbook` interface)."""
    for attr_name in dir(mod):
        try:
            attr = getattr(mod, attr_name)
        except (AttributeError, OSError):
            continue

        if not hasattr(attr, "_reg_clsid_"):
            continue

        coclass_name = str(attr_name)
        target = model["classes"].setdefault(
            coclass_name, {"type": "Class", "members": {}},
        )

        com_interfaces = getattr(attr, "_com_interfaces_", None)
        if not com_interfaces:
            continue
        default_intf = com_interfaces[0]
        intf_name = getattr(default_intf, "__name__", None)
        if not intf_name or intf_name not in model["classes"]:
            continue

        src = model["classes"][intf_name]["members"]
        target["members"].update(src)


# ---------------------------------------------------------------- driver


def generate_model(
    references_path: Path,
    output_path: Path,
    *,
    promote_application: bool = True,
) -> dict[str, Any]:
    """Drive both passes and return the resulting model dict."""
    cc = _import_comtypes()

    LOG.info("Loading references from %s", references_path)
    with open(references_path, "r", encoding="utf-8") as fh:
        ref_data = json.load(fh)

    references = ref_data.get("references", [])
    if not references:
        sys.stderr.write(
            "Error: vba_references.json contained no `references` array. "
            "Did the VBA exporter complete successfully?\n"
        )
        sys.exit(2)

    model: dict[str, Any] = {
        "globals": {},
        "classes": {},
        "enums": {},
        "references": references,
    }

    for ref in references:
        ref_name = ref.get("name", "Unknown")
        LOG.info("Pass 1 — interfaces / enums: %s", ref_name)
        mod = _load_module(cc, ref)
        if mod is None:
            LOG.warning("  could not load %s — skipped", ref_name)
            continue
        _process_interfaces_and_enums(model, mod, ref_name)

    for ref in references:
        ref_name = ref.get("name", "Unknown")
        LOG.info("Pass 2 — coclasses: %s", ref_name)
        mod = _load_module(cc, ref)
        if mod is None:
            continue
        _process_coclasses(model, mod)

    if promote_application:
        _promote_application_globals(model)

    LOG.info("Writing %s", output_path)
    output_path.write_text(json.dumps(model, indent=2), encoding="utf-8")
    LOG.info(
        "Done — %d globals, %d classes, %d enums.",
        len(model["globals"]), len(model["classes"]), len(model["enums"]),
    )
    return model


def _promote_application_globals(model: dict) -> None:
    """Lift every member of the host's `Application` class into the
    global scope (so `ActiveWindow`, `ActiveDocument`, … resolve without
    qualification). Tries common host application interfaces."""
    candidates = [
        "Application", "_Application",
        "IVApplication",      # Visio
        "IAcadApplication",   # AutoCAD
    ]
    for cls_name in candidates:
        cls = model["classes"].get(cls_name)
        if not cls:
            continue
        for m_name, m_def in cls["members"].items():
            model["globals"].setdefault(m_name, m_def)
        return  # only promote the first hit


def _build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="generate_model.py",
        description=(
            "Generate a VBAlidator JSON object model from the COM type "
            "libraries listed in vba_references.json."
        ),
    )
    p.add_argument(
        "references",
        nargs="?",
        default="vba_references.json",
        help="Path to the vba_references.json produced by VBA_Model_Exporter.bas. "
             "Defaults to ./vba_references.json or ./tools/vba_references.json.",
    )
    p.add_argument(
        "-o", "--output",
        default="vba_model.json",
        help="Where to write the resulting model (default: vba_model.json).",
    )
    p.add_argument(
        "--no-app-promote",
        action="store_true",
        help="Do not lift Application members into global scope. Use this "
             "when you intend to qualify every host call (e.g. "
             "`Application.ActiveDocument`).",
    )
    p.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Verbose logging (per-library progress).",
    )
    return p


def _resolve_references_arg(arg: str) -> Path:
    p = Path(arg)
    if p.is_file():
        return p
    # Fall back to a couple of conventional locations.
    for c in (
        Path("vba_references.json"),
        Path("tools") / "vba_references.json",
        Path("..") / "vba_references.json",
    ):
        if c.is_file():
            return c.resolve()
    sys.stderr.write(
        f"Error: could not find {arg!r} (or vba_references.json in the "
        "usual spots). Run VBA_Model_Exporter.bas first.\n"
    )
    sys.exit(2)


def main() -> None:
    args = _build_argparser().parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
    )
    refs = _resolve_references_arg(args.references)
    out = Path(args.output)
    generate_model(refs, out, promote_application=not args.no_app_promote)


if __name__ == "__main__":
    main()
