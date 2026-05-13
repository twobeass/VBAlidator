#!/usr/bin/env python3
"""Build `src/models/mscomctl.json` from the system-installed
`MSCOMCTL.OCX`.

The COM type library only exposes the control-specific properties
(`Nodes`, `Indentation`, `Style`, …). The VB6 host injects a base set
of *container-control* members (`Move`, `Left`, `Top`, `Width`, …) onto
every visible control at runtime — those don't live in the OCX and so
don't show up in the generated model. This script:

1. Runs the standard `tools/generate_model.py` over the OCX.
2. Patches each user-facing control class (`TreeView`, `ListView`,
   `Toolbar`, …) with the VB6 container-control member set.

Output: `src/models/mscomctl.json`.

Platform: Windows + the OCX registered (typical for any machine that
ever had VB6 / classic Office).
"""
from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path

from generate_model import generate_model

LOG = logging.getLogger("vbalidator.build_mscomctl")

# Discovered control-class names that get the VB6 base-member injection.
# Restricted to user-facing visual controls; collections / data items
# (Node, ListItem, Buttons, Tabs, …) are excluded because they aren't
# child controls of a form and don't carry container-control members.
CONTROL_CLASSES = (
    "TreeView", "TreeView2", "TreeView3", "TreeView4",
    "ListView", "ListView2", "ListView3", "ListView4",
    "Toolbar", "Toolbar2", "Toolbar3",
    "ProgressBar", "ProgressBar2",
    "Slider", "Slider2",
    "StatusBar", "StatusBar2", "StatusBar3",
    "TabStrip", "TabStrip2", "TabStrip3",
    "ImageList", "ImageList2", "ImageList3",
    "ImageCombo", "ImageCombo2", "ImageCombo3",
    "Animation", "MonthView", "DTPicker", "FlatScrollBar", "UpDown",
)

# VB6 container-control base members. `Move` is the one that drives the
# JSONBag fixture; the rest are the surface most resize / layout code in
# real VBA forms reaches for. Types from the VB6 OLE Control interface
# (IOleControl) — Long for size/position, Boolean for visibility flags,
# Variant for object references.
BASE_CONTROL_MEMBERS = {
    "Move": {"type": "Sub", "min_args": 1, "max_args": 4},
    "Refresh": {"type": "Sub", "min_args": 0, "max_args": 0},
    "SetFocus": {"type": "Sub", "min_args": 0, "max_args": 0},
    "ZOrder": {"type": "Sub", "min_args": 0, "max_args": 1},
    "ShowWhatsThis": {"type": "Sub", "min_args": 0, "max_args": 0},
    "Drag": {"type": "Sub", "min_args": 0, "max_args": 1},
    "Left": {"type": "Long"},
    "Top": {"type": "Long"},
    "Width": {"type": "Long"},
    "Height": {"type": "Long"},
    "Visible": {"type": "Boolean"},
    "Enabled": {"type": "Boolean"},
    "TabIndex": {"type": "Integer"},
    "TabStop": {"type": "Boolean"},
    "Tag": {"type": "String"},
    "Index": {"type": "Integer"},
    "Name": {"type": "String"},
    "Container": {"type": "Object"},
    "Parent": {"type": "Object"},
    "HelpContextID": {"type": "Long"},
    "WhatsThisHelpID": {"type": "Long"},
    "Object": {"type": "Object"},
    "DragMode": {"type": "Integer"},
    "DragIcon": {"type": "Object"},
    "CausesValidation": {"type": "Boolean"},
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


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    repo_root = Path(__file__).resolve().parent.parent

    # Default OCX path; override with `MSCOMCTL_OCX` env var if needed.
    default_ocx = Path(
        os.environ.get(
            "MSCOMCTL_OCX",
            r"C:\Program Files\Microsoft Office\root\vfs\System\MSCOMCTL.OCX",
        )
    )
    if not default_ocx.is_file():
        sys.stderr.write(
            f"MSCOMCTL.OCX not found at {default_ocx}. Set the MSCOMCTL_OCX "
            f"environment variable to the absolute path of the OCX.\n"
        )
        return 1

    # Write a temporary refs file and drive generate_model with it.
    refs_path = repo_root / "_mscomctl_refs.json"
    refs_path.write_text(json.dumps({
        "references": [
            {"name": "MSComctlLib", "path": str(default_ocx)},
        ],
    }), encoding="utf-8")

    out_path = repo_root / "src" / "models" / "mscomctl.json"
    LOG.info("Generating from %s → %s", default_ocx, out_path)
    model = generate_model(refs_path, out_path)

    n = _patch_controls(model)
    LOG.info("Injected VB6 container-control members into %d control classes", n)

    # Write the final model. generate_model already wrote one; we overwrite.
    out_path.write_text(json.dumps(model, indent=2), encoding="utf-8")
    LOG.info(
        "Done — %d classes, %d globals, %d enums (file: %d KB)",
        len(model["classes"]),
        len(model["globals"]),
        len(model["enums"]),
        out_path.stat().st_size // 1024,
    )

    # Clean up the scratch refs file.
    try:
        refs_path.unlink()
    except OSError:
        pass

    return 0


if __name__ == "__main__":
    sys.exit(main())
