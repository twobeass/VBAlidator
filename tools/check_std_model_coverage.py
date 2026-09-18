#!/usr/bin/env python3
"""Report VBA intrinsics and constants that `src/std_model.json` is missing.

std_model.json is the model every host layers on top of, so a name absent from
it is a false "undefined identifier" for every user of every host. It is
hand-maintained, which means gaps are found the way all gaps in hand-maintained
lists are found: one user at a time, by tripping over them.

This checks it against Microsoft's own VBA Language Reference instead. That
reference is published from MicrosoftDocs/VBA-Docs, one page per intrinsic:

    Language/Reference/User-Interface-Help/abs-function.md
    Language/Reference/User-Interface-Help/beep-statement.md
    Language/Reference/User-Interface-Help/colour-constants.md

    usage: check_std_model_coverage.py <path to a VBA-Docs clone>

To get the clone without downloading all of Office:

    git clone --filter=blob:none --no-checkout --depth 1 \\
        https://github.com/MicrosoftDocs/VBA-Docs.git
    cd VBA-Docs
    git sparse-checkout init --no-cone
    git sparse-checkout set '/Language'
    git checkout

STATEMENTS ARE NOT REPORTED. Print, Input, Close, Open, Get, Put, Line and the
rest are handled as statement keywords in the analyser, not resolved as
identifiers, so adding them here would fight that handling rather than help.
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HELP = os.path.join("Language", "Reference", "User-Interface-Help")

# Handled as statement keywords by the analyser - see the module docstring.
STATEMENT_KEYWORDS = set("""
    error input print close open get put let stop len mid lset rset erase
    redim line write seek lock unlock name kill width spc tab
""".split())


def published(clone):
    """Every intrinsic function name, and every vb* constant, that the VBA
    Language Reference documents."""
    d = os.path.join(clone, HELP)
    if not os.path.isdir(d):
        sys.exit(f"no {HELP} under {clone} - see the usage note in this file")

    functions, constants = {}, set()
    for fn in sorted(os.listdir(d)):
        if not fn.endswith(".md"):
            continue
        text = open(os.path.join(d, fn), encoding="utf-8", errors="replace").read()

        if fn.endswith("-function.md"):
            m = re.search(r"^#\s+([A-Za-z0-9_$]+)\s+function", text, re.M | re.I)
            if m:
                rt = re.search(
                    r"[Rr]eturn\s+(?:value\s+)?(?:type\s+)?\*\*([A-Za-z]+)\*\*", text)
                # Variant when the page doesn't state one: permissive, and the
                # alternative is inventing a type.
                functions[m.group(1)] = rt.group(1) if rt else "Variant"
        elif "constants" in fn:
            constants |= set(re.findall(r"\*\*(vb[A-Za-z0-9_]+)\*\*", text))

    return functions, constants


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__.strip().splitlines()[-1])

    functions, constants = published(sys.argv[1])
    have = json.load(open(os.path.join(ROOT, "src", "std_model.json")))["globals"]

    miss_fn = {n: t for n, t in functions.items()
               if n not in have and n.lower() not in STATEMENT_KEYWORDS}
    miss_const = sorted(n for n in constants if n not in have)

    print(f"VBA Language Reference: {len(functions)} functions, "
          f"{len(constants)} vb* constants")
    print(f"src/std_model.json:     {len(have)} globals")
    print()
    print(f"missing functions ({len(miss_fn)}):")
    for n in sorted(miss_fn):
        print(f'    "{n}": {{ "type": "Function", "returns": "{miss_fn[n]}" }},')
    print()
    print(f"missing vb* constants ({len(miss_const)}):")
    for n in miss_const:
        print(f'    "{n}": {{ "type": "Long" }},')

    return 1 if (miss_fn or miss_const) else 0


if __name__ == "__main__":
    sys.exit(main())
