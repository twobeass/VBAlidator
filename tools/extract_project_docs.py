#!/usr/bin/env python3
"""Extract the MS Project object model from MicrosoftDocs/VBA-Docs.

That repository is the SOURCE of learn.microsoft.com/office/vba/api/project.* -
the published pages are built from these markdown files - so parsing it gives
the documented object model without a single name being typed by hand.

Which matters, because a model is a WHITELIST: a name listed wrongly makes the
analyser bless a typo, which is worse than having no model. A hand-transcribed
first pass dropped Resource.UniqueID and Assignment.ResourceUniqueID and
invented Resource.Generic, Resource.Inactive and nine pj* constants. None of
those mistakes are possible here.

It also sidesteps a real documentation bug. The project.assignments OBJECT page
carries the Assignment object's member tables by mistake, every link pointing at
project.assignment.* - read that page and you would swear the collection had
Work, Cost and Text1..30. But there is no Project.Assignments.Work.md FILE, and
this reads the member files. Six members, which is the truth.

    usage: extract_project_docs.py <VBA-Docs clone> [-d tools/data]

Writes both files build_project_model.py needs:

    project_api.json         classes and their members
    project_vba_enums.json   the 172 enumerations and their 4,907 pj* constants

To get the clone without downloading all of Office (VBA-Docs covers every app):

    git clone --filter=blob:none --no-checkout --depth 1 \\
        https://github.com/MicrosoftDocs/VBA-Docs.git
    cd VBA-Docs
    git sparse-checkout init --no-cone
    git sparse-checkout set '/api/project.*' '/project'
    git checkout

The extracted JSON is committed, so building the model afterwards needs neither
the clone nor the network.

WHAT THIS CANNOT SEE. Documentation describes the SUPPORTED API. A type library
also carries hidden and legacy members, and Project.NewTasksAreManual is one:
undocumented, absent from every page here, and real - code in the wild assigns
to it from an early-bound Project variable, which could not compile otherwise.
So this is authoritative for what is SUPPORTED, not for what exists; the
build script carries such members as explicit exceptions.
"""
import argparse
import json
import os
import re
import sys
from collections import defaultdict

# "# Task.UniqueID property (Project)" / "# Tasks object (Project)"
H1 = re.compile(r"^#\s+(?P<name>[A-Za-z0-9_.]+)\s+"
                r"(?P<kind>property|method|event|object|enumeration)\s+\(Project\)",
                re.M)
# "Read-only **Long**." / "Read/write **Variant**."
PROPTYPE = re.compile(r"Read(?:-only|/write)\s+\*\*([A-Za-z0-9_]+)\*\*")
RETURN = re.compile(r"^##\s*Return value\s*$(?P<body>.*?)(?=^##|\Z)", re.M | re.S)
# One row of the Parameters table:
#     | _StartDate_|Required|**Variant**|The start date ... |
PARAM = re.compile(
    r"^\|\s*_(?P<pname>\w+)_\s*\|\s*(?P<req>Required|Optional)\s*\|"
    r"\s*\*{0,2}(?P<ptype>[A-Za-z0-9_]+)?",
    re.M | re.I)
# One row of an enumeration's Name/Value table:
#     |**pjResourceTypeWork**|0|Resource type is **Work**.|
# Five pages bold the value too - |**pjResourceWarningEngagementViolation**|**8**||
# - which is why the value's asterisks are optional rather than assumed absent.
# and one page writes the value as "32 (&H20)", decimal then hex, so anything
# after the number is skipped rather than assumed to be the end of the cell.
ENUMROW = re.compile(
    r"^\|\s*\*\*(?P<cname>\w+)\*\*\s*\|\s*\*{0,2}(?P<val>-?\w+)\*{0,2}[^|]*\|", re.M)


def parse(path):
    text = open(path, encoding="utf-8", errors="replace").read()
    m = H1.search(text)
    if not m:
        return None

    name, kind = m.group("name"), m.group("kind")
    out = {"kind": kind, "file": os.path.basename(path)}

    if kind == "property":
        t = PROPTYPE.search(text)
        out["type"] = t.group(1) if t else None
        # Properties CAN take arguments - Project.BaselineSavedDate does. Not
        # assuming zero is the difference between checking a call and rejecting
        # a valid one.
        out["params"] = params(text)
    elif kind == "method":
        r = RETURN.search(text)
        if r:
            rt = re.search(r"\*\*([A-Za-z0-9_]+)\*\*", r.group("body"))
            body = r.group("body").strip()
            out["type"] = rt.group(1) if rt else ("Nothing" if "Nothing" in body else None)
        else:
            # No Return value section at all is how MS writes a Sub.
            out["type"] = "Nothing"
        out["params"] = params(text)
    return name, out


def enum_members(text):
    """Name -> value from an enumeration page's table. Values are decimal in
    every current page; anything that isn't an integer is kept verbatim rather
    than coerced, so a hex or symbolic value would survive rather than become
    a wrong number."""
    out = {}
    for m in ENUMROW.finditer(text):
        raw = m.group("val")
        try:
            out[m.group("cname")] = int(raw)
        except ValueError:
            out[m.group("cname")] = raw
    return out


def params(text):
    """Per-argument name, optionality and documented type, in declaration order.

    Order matters and counts alone lose it. Application.FilterEdit has required
    arguments at positions 1, 2 and 9 with optional ones in between, so
    "3 required of 13" accepts three positional arguments that Project would
    reject. Keeping the list lets the model say which positions those are, and
    carries the argument NAMES - which is how Project's command methods are
    actually called (FilterEdit Name:="x", Test:="contains"), so it is the
    difference between checking those calls and not."""
    out = []
    for m in PARAM.finditer(text):
        out.append({
            "name": m.group("pname"),
            "optional": m.group("req").lower() == "optional",
            "type": m.group("ptype") or None,
        })
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("clone", help="path to a VBA-Docs clone")
    ap.add_argument("-d", "--data-dir", default=None,
                    help="where to write both JSON files (default: tools/data)")
    args = ap.parse_args()

    api = os.path.join(args.clone, "api")
    if not os.path.isdir(api):
        sys.exit(f"no api/ directory under {args.clone}")

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = args.data_dir or os.path.join(root, "tools", "data")
    out_path = os.path.join(data_dir, "project_api.json")
    enum_path = os.path.join(data_dir, "project_vba_enums.json")

    classes = defaultdict(dict)
    objects, events, skipped, enums = {}, defaultdict(list), [], {}

    for fn in sorted(os.listdir(api)):
        if not fn.lower().startswith("project.") or not fn.endswith(".md"):
            continue
        got = parse(os.path.join(api, fn))
        if not got:
            skipped.append(fn)
            continue
        name, info = got

        if info["kind"] == "enumeration":
            members = enum_members(
                open(os.path.join(api, fn), encoding="utf-8", errors="replace").read())
            if members:
                enums[name] = {
                    "members": members,
                    "url": "https://learn.microsoft.com/en-us/office/vba/api/"
                           + fn[:-3].lower(),
                }
            continue
        if info["kind"] == "object":
            # The H1 carries the canonical casing; filenames are inconsistent
            # about it (both Project.Application.md and Project.application.md
            # exist).
            objects[name.lower()] = name
            continue

        if "." not in name:
            skipped.append(fn)
            continue
        cls, member = name.rsplit(".", 1)
        if info["kind"] == "event":
            # Events are not callable members - modelling them as methods would
            # make ActiveProject.Calculate look like a valid call.
            events[cls].append(member)
            continue
        classes[cls][member] = {k: v for k, v in info.items() if k != "kind"} | {
            "kind": info["kind"]}

    data = {
        "source": "github.com/MicrosoftDocs/VBA-Docs (api/project.*)",
        "objects": objects,
        "classes": {c: classes[c] for c in sorted(classes)},
        "events": {c: sorted(v) for c, v in sorted(events.items())},
    }
    os.makedirs(data_dir, exist_ok=True)
    with open(out_path, "w") as fh:
        json.dump(data, fh, indent=1, sort_keys=True)
        fh.write("\n")
    with open(enum_path, "w") as fh:
        json.dump(enums, fh, indent=1, sort_keys=True)
        fh.write("\n")

    total = sum(len(m) for m in classes.values())
    print(f"wrote {out_path}")
    print(f"  {len(classes)} classes with members, {total} members")
    print(f"wrote {enum_path}")
    print(f"  {len(enums)} enumerations, "
          f"{sum(len(e['members']) for e in enums.values())} constants")
    print(f"  {len(objects)} object pages, "
          f"{sum(len(v) for v in events.values())} events (not modelled)")
    if skipped:
        print(f"  {len(skipped)} files with no parseable H1, e.g. {skipped[:3]}")
    big = sorted(classes.items(), key=lambda kv: -len(kv[1]))[:8]
    for c, m in big:
        props = sum(1 for v in m.values() if v["kind"] == "property")
        print(f"    {c:14} {len(m):4}  ({props} properties, {len(m)-props} methods)")
    missing = [f"{c}.{n}" for c, m in classes.items()
               for n, v in m.items() if not v.get("type")]
    if missing:
        print(f"  {len(missing)} members whose page states no type "
              f"(kept as Variant by the builder): {missing[:5]}")


if __name__ == "__main__":
    main()
