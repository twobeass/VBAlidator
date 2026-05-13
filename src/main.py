import argparse
import json
import os
import sys

from colorama import Fore, Style, init

from . import __version__
from .api import precheck

init(autoreset=True)


def _color_for_severity(sev):
    return {
        "error": Fore.RED,
        "compile_verified": Fore.RED,
        "warning": Fore.YELLOW,
        "info": Fore.CYAN,
    }.get(sev, Fore.WHITE)


def _color_for_score(score):
    if score >= 90:
        return Fore.GREEN
    if score >= 70:
        return Fore.YELLOW
    return Fore.RED


def _emit(quiet, *args, **kwargs):
    """Print to stdout only when `--quiet` is off. Errors should use the
    standard `print(..., file=sys.stderr)` instead so they survive
    `--quiet`."""
    if not quiet:
        print(*args, **kwargs)


def main():
    parser = argparse.ArgumentParser(
        description="VBAlidator — VBA static analyser & compile-safety prechecker"
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"vbalidator {__version__}",
        help="Print the package version and exit.",
    )
    parser.add_argument(
        "input_path",
        help="Path to a VBA file (.bas/.cls/.frm) or a folder containing them.",
    )
    parser.add_argument(
        "--define",
        help="Conditional compilation constants, e.g. 'WIN64=True,VBA7=True'",
    )
    parser.add_argument(
        "--model",
        help="Path to a custom JSON object model definition file.",
    )
    parser.add_argument(
        "--host",
        choices=[
            "excel", "word", "access", "outlook", "visio",
            "mscomctl", "msforms",
            "scripting", "vbscript_regexp", "wscript_shell", "shell_application",
        ],
        help="Built-in host model to load (Excel/Word/Access/Outlook/Visio/"
             "MSComCtl/MSForms/Scripting/VBScript_RegExp/WScript_Shell/"
             "Shell_Application). Bundled models/<host>.json is layered on "
             "top of the standard model so the user does not need to run "
             "the VBA_Model_Exporter.bas first. Excel/Word/Access/Visio are "
             "full-fidelity models from real Office type libraries; the "
             "companion stubs (MSComCtl/MSForms/Scripting/…) cover the "
             "common COM libraries and **auto-layer** whenever any scanned "
             "file mentions their ProgID / namespace (`Scripting.Dictionary`, "
             "`VBScript.RegExp`, `WScript.Shell`, `Shell.Application`, "
             "`MSForms.X`, or a `.frm` referencing ComctlLib) — explicit "
             "`--host <name>` is rarely needed. Outlook is a minimal "
             "hand-curated stub (the COM/TLB path is GPO-blocked on most "
             "managed installs).",
    )
    parser.add_argument(
        "--output",
        default="vba_report.json",
        help="Path to write the JSON v2 report (default: vba_report.json).",
    )
    parser.add_argument(
        "--score-threshold",
        type=int,
        default=90,
        help="Minimum confidence score to consider the input compile-safe "
             "(default 90). The CLI exits non-zero when the score is below "
             "this value or when there is at least one severity=error issue.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        default=True,
        help="Count severity=warning toward the gating set (default).",
    )
    parser.add_argument(
        "--no-strict",
        dest="strict",
        action="store_false",
        help="Ignore warnings/info for gating; only errors block.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress all non-essential stdout (banner, per-issue list, "
             "summary). Errors continue to be written to stderr and the "
             "exit code is unaffected. The JSON report is still written "
             "to --output.",
    )
    parser.add_argument(
        "--roundtrip",
        action="store_true",
        help="Drive the actual VBE compiler through Office COM as a "
             "second-opinion check (Windows + Office only). Each compile "
             "error returned by VBE is reported with severity "
             "'compile_verified'. Falls back to a single info-level "
             "notice when the platform / Python bindings are missing.",
    )

    args = parser.parse_args()

    if not os.path.exists(args.input_path):
        # Hard errors always go to stderr; --quiet must not hide them.
        print(
            Fore.RED + f"Error: input path '{args.input_path}' does not exist.",
            file=sys.stderr,
        )
        sys.exit(2)

    defines = {}
    if args.define:
        for pair in args.define.split(","):
            if "=" in pair:
                k, v = pair.split("=", 1)
                low = v.strip().lower()
                if low == "true":
                    defines[k.strip().upper()] = True
                elif low == "false":
                    defines[k.strip().upper()] = False
                else:
                    defines[k.strip().upper()] = v.strip()

    _emit(args.quiet, Fore.CYAN + f"VBAlidator: scanning {args.input_path}"
          + (f" (host={args.host})" if args.host else ""))

    try:
        result = precheck(
            args.input_path,
            host=args.host,
            model_path=args.model,
            defines=defines,
            strict=args.strict,
            roundtrip=args.roundtrip,
        )
    except Exception as exc:  # surface unexpected pipeline failures
        print(Fore.RED + f"Pipeline error: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(3)

    # Per-issue + summary console output (suppressed under --quiet).
    if not args.quiet:
        for issue in result.issues:
            sev_color = _color_for_severity(issue.get("severity", "error"))
            print(
                f"{Fore.MAGENTA}{issue.get('file','')}:{issue.get('line',0)}: "
                f"{sev_color}{issue.get('severity','error').upper()}{Style.RESET_ALL}  "
                f"{Fore.WHITE}[{issue.get('rule_id','VBA000')}]  "
                f"{issue.get('message','')}"
            )

        s = result.json()["summary"]
        score_color = _color_for_score(result.score)
        print()
        print(f"{Fore.CYAN}Files scanned : {Style.RESET_ALL}{s['files_scanned']}")
        print(f"{Fore.CYAN}Errors        : {Fore.RED}{s['errors']}")
        print(f"{Fore.CYAN}Warnings      : {Fore.YELLOW}{s['warnings']}")
        print(f"{Fore.CYAN}Info          : {Fore.WHITE}{s['info']}")
        print(f"{Fore.CYAN}Confidence    : {score_color}{result.score} / 100"
              f"  {'(compile-safe)' if result.compile_safe else '(needs fixes)'}")

    # JSON output is always written — the file is the primary product, the
    # stdout summary is decorative. Failure to write the file is a hard
    # error and goes to stderr.
    try:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result.json(), f, indent=2)
        _emit(args.quiet, f"{Fore.CYAN}Report saved  : {Style.RESET_ALL}{args.output}")
    except OSError as exc:
        print(Fore.RED + f"Could not write report: {exc}", file=sys.stderr)
        sys.exit(4)

    # Exit code: 0 only when score >= threshold and no errors.
    blocking = (not result.compile_safe) or (result.score < args.score_threshold)
    sys.exit(1 if blocking else 0)


if __name__ == "__main__":
    main()
