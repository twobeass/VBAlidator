import argparse
import json
import os
import sys
from colorama import init, Fore, Style

from .api import precheck

init(autoreset=True)


def _color_for_severity(sev):
    return {
        "error": Fore.RED,
        "warning": Fore.YELLOW,
        "info": Fore.CYAN,
    }.get(sev, Fore.WHITE)


def _color_for_score(score):
    if score >= 90:
        return Fore.GREEN
    if score >= 70:
        return Fore.YELLOW
    return Fore.RED


def main():
    parser = argparse.ArgumentParser(
        description="VBAlidator — VBA static analyser & compile-safety prechecker"
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
        choices=["excel", "word", "access", "outlook"],
        help="Built-in host model to load (Excel/Word/Access/Outlook). "
             "When set, the bundled models/<host>.json is layered on top of "
             "the standard model so the user does not need to run the "
             "VBA_Model_Exporter.bas first.",
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
        help="Suppress per-issue console output. Only print summary + score.",
    )

    args = parser.parse_args()

    if not os.path.exists(args.input_path):
        print(Fore.RED + f"Error: input path '{args.input_path}' does not exist.")
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

    print(Fore.CYAN + f"VBAlidator: scanning {args.input_path}"
          + (f" (host={args.host})" if args.host else ""))

    try:
        result = precheck(
            args.input_path,
            host=args.host,
            model_path=args.model,
            defines=defines,
            strict=args.strict,
        )
    except Exception as exc:  # surface unexpected pipeline failures
        print(Fore.RED + f"Pipeline error: {exc}")
        import traceback
        traceback.print_exc()
        sys.exit(3)

    # Console output
    if not args.quiet:
        for issue in result.issues:
            sev_color = _color_for_severity(issue.get("severity", "error"))
            print(
                f"{Fore.MAGENTA}{issue.get('file','')}:{issue.get('line',0)}: "
                f"{sev_color}{issue.get('severity','error').upper()}{Style.RESET_ALL}  "
                f"{Fore.WHITE}[{issue.get('rule_id','VBA000')}]  "
                f"{issue.get('message','')}"
            )

    # Summary block
    s = result.json()["summary"]
    score_color = _color_for_score(result.score)
    print()
    print(f"{Fore.CYAN}Files scanned : {Style.RESET_ALL}{s['files_scanned']}")
    print(f"{Fore.CYAN}Errors        : {Fore.RED}{s['errors']}")
    print(f"{Fore.CYAN}Warnings      : {Fore.YELLOW}{s['warnings']}")
    print(f"{Fore.CYAN}Info          : {Fore.WHITE}{s['info']}")
    print(f"{Fore.CYAN}Confidence    : {score_color}{result.score} / 100"
          f"  {'(compile-safe)' if result.compile_safe else '(needs fixes)'}")

    # JSON output
    try:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result.json(), f, indent=2)
        print(f"{Fore.CYAN}Report saved  : {Style.RESET_ALL}{args.output}")
    except OSError as exc:
        print(Fore.RED + f"Could not write report: {exc}")
        sys.exit(4)

    # Exit code: 0 only when score >= threshold and no errors.
    blocking = (not result.compile_safe) or (result.score < args.score_threshold)
    sys.exit(1 if blocking else 0)


if __name__ == "__main__":
    main()
