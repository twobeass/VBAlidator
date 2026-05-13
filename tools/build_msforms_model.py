#!/usr/bin/env python3
"""Build `src/models/msforms.json` from the system-installed FM20.DLL
(Microsoft Forms 2.0 Object Library).

Two reasons we don't just consume the comtypes extract verbatim:

1. **`MSForms` namespace prefix.** stdVBA-style code uses
   `TypeOf ctrl Is MSForms.UserForm` patterns — `MSForms` needs to
   resolve as a global namespace so the analyzer doesn't flag
   `Undefined identifier 'MSForms'`. We register it as a `Variant`
   global (same trick the std_model already uses for `Conversion`,
   `Math`, `Strings`, … to make member access go through permissive).

2. **VB6 container-control members.** The FM20 type-library only
   carries the per-control properties; `Move`, `Refresh`, `Left/Top/
   Width/Height`, `Visible/Enabled`, `Tag/Name/Container/Parent` live
   on the VB6 OLE control host and are inherited at runtime. We patch
   them onto every visible-control class so resize/layout code in real
   UserForm modules type-checks.

Output: `src/models/msforms.json`.

Platform: Windows + FM20.DLL registered (default for any machine that
ever had Office VBA).
"""
from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path

from generate_model import generate_model

LOG = logging.getLogger("vbalidator.build_msforms")

# Visible-control classes that pick up the VB6 base-member injection.
# Collections and helpers (Controls, Frames, Pages, …) are excluded —
# they aren't form children and don't carry container-control members.
CONTROL_CLASSES = (
    "UserForm", "Frame", "MultiPage", "TabStrip",
    "Label", "Image",
    "TextBox", "ComboBox", "ListBox",
    "CheckBox", "OptionButton", "ToggleButton",
    "CommandButton",
    "ScrollBar", "SpinButton",
    # `Control` is the abstract base — adding it here gives library code
    # like `this.Control.Caption` / `.Left` a place to resolve members
    # when the variable is typed as the base (stdVBA does this in
    # stdUIElement.cls).
    "Control",
)

# VB6 container-control base members — the runtime injection any visible
# control on a UserForm picks up regardless of what the OCX exposes.
# Same shape as the MSComCtl variant so future readers spot the pattern.
BASE_CONTROL_MEMBERS = {
    "Move": {"type": "Sub", "min_args": 1, "max_args": 4},
    "Refresh": {"type": "Sub", "min_args": 0, "max_args": 0},
    "SetFocus": {"type": "Sub", "min_args": 0, "max_args": 0},
    "ZOrder": {"type": "Sub", "min_args": 0, "max_args": 1},
    "Left": {"type": "Single"},
    "Top": {"type": "Single"},
    "Width": {"type": "Single"},
    "Height": {"type": "Single"},
    "Visible": {"type": "Boolean"},
    "Enabled": {"type": "Boolean"},
    "TabIndex": {"type": "Integer"},
    "TabStop": {"type": "Boolean"},
    "Tag": {"type": "String"},
    "Name": {"type": "String"},
    "Container": {"type": "Object"},
    "Parent": {"type": "Object"},
    "HelpContextID": {"type": "Long"},
    "Object": {"type": "Object"},
    "ControlSource": {"type": "String"},
    "ControlTipText": {"type": "String"},
    "Accelerator": {"type": "String"},
    "BackColor": {"type": "Long"},
    "ForeColor": {"type": "Long"},
    # `Caption` / `Value` aren't on every visible control (TextBox uses
    # `Text`, ScrollBar uses `Value`, …) but they're on enough of them
    # that library code reaches for them via the abstract `Control`
    # base. Adding them here is pragmatic over precise; the alternative
    # is a per-subclass switch the model schema doesn't support yet.
    "Caption": {"type": "String"},
    "Value": {"type": "Variant"},
}


def _patch_controls(model: dict) -> int:
    """Inject VB6 base control members into every shipped control class.
    Returns the number of classes patched."""
    patched = 0
    for cls_name in CONTROL_CLASSES:
        cls = model["classes"].get(cls_name)
        if cls is None:
            continue
        members = cls.setdefault("members", {})
        for name, defn in BASE_CONTROL_MEMBERS.items():
            members.setdefault(name, dict(defn))
        patched += 1
    return patched


def _register_namespace(model: dict) -> None:
    """`TypeOf ctrl Is MSForms.UserForm` and `Dim x As MSForms.TextBox`
    need the `MSForms` prefix to resolve as something — anything — that
    permits member access. `Variant` is the same trick std_model uses
    for the VBA library namespaces (`Conversion`, `Math`, `Strings`)."""
    model["globals"]["MSForms"] = {"type": "Variant"}


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    repo_root = Path(__file__).resolve().parent.parent

    default_dll = Path(
        os.environ.get(
            "FM20_DLL",
            r"C:\Program Files\Microsoft Office\root\vfs\System\FM20.DLL",
        )
    )
    if not default_dll.is_file():
        sys.stderr.write(
            f"FM20.DLL not found at {default_dll}. Set the FM20_DLL "
            f"environment variable to the absolute path of the DLL.\n"
        )
        return 1

    refs_path = repo_root / "_msforms_refs.json"
    refs_path.write_text(json.dumps({
        "references": [
            {"name": "MSForms", "path": str(default_dll)},
        ],
    }), encoding="utf-8")

    out_path = repo_root / "src" / "models" / "msforms.json"
    LOG.info("Generating from %s → %s", default_dll, out_path)
    model = generate_model(refs_path, out_path)

    n = _patch_controls(model)
    _register_namespace(model)
    LOG.info("Injected VB6 container-control members into %d control classes", n)
    LOG.info("Registered `MSForms` as a global namespace alias")

    out_path.write_text(json.dumps(model, indent=2), encoding="utf-8")
    LOG.info(
        "Done — %d classes, %d globals, %d enums (file: %d KB)",
        len(model["classes"]),
        len(model["globals"]),
        len(model["enums"]),
        out_path.stat().st_size // 1024,
    )

    try:
        refs_path.unlink()
    except OSError:
        pass

    return 0


if __name__ == "__main__":
    sys.exit(main())
