# TODO — what still needs human / Windows / live-account access

The Phase 0–5 work in PR #14 is complete from the analyser's side, but a
handful of items can only be done with **a real Windows + Office host**,
**live PyPI / GHCR / GitHub-Pages accounts**, or **deliberate human
review**. They are listed here so they don't get lost, grouped by who can
do them.

Status legend: `[ ]` open · `[~]` partially done in the sandbox · `[x]` done.


## A. Things that need a Windows + Office machine

These cannot run on the Linux GitHub-hosted runners.

- [x] **Regenerate every bundled host model from a real Office install.**
  Shipped in PR #23 (v1.1.0) — Excel/Word/Access/Visio are now full-
  fidelity exports of the real Office type libraries (1–3 MB each).
  Plan C ("opt-in `*-full` variants") was abandoned because shipping
  the full models in the default install turned out to be the cleaner
  default. Outlook remains a hand-curated stub (the Trust-Center
  AccessVBOM path is GPO-blocked on managed installs — see
  [memory/office-quirks.md] for context).

  Companion COM stubs added in iter-6 / iter-7:
  - `src/models/mscomctl.json` (PR #28) — Microsoft Common Controls
  - `src/models/msforms.json` (PR #30) — MSForms 2.0 UserForm controls
  - `src/models/scripting.json` (PR #40) — Scripting.Dictionary / FSO
  - `src/models/vbscript_regexp.json` (PR #40)
  - `src/models/wscript_shell.json` (PR #40)
  - `src/models/shell_application.json` (PR #40)

  Regen scripts (`tools/build_mscomctl_model.py` / `build_msforms_model.py`)
  are deterministic re-runs of the comtypes extract + VB6 container-
  control patch. The four scripting/COM stubs are hand-curated.

  Manual regen workflow (still valid for project-specific custom
  references):
    1. Open Excel → VBE → File ▶ Import File → `tools/VBA_Model_Exporter.bas`.
    2. Run `ExportReferences`. It writes `vba_references.json`.
    3. `pip install comtypes` → `python tools/generate_model.py vba_references.json -o vba_model.json`.
    4. `vbalidator … --model vba_model.json`.

- [~] **Run the VBE round-trip suite end-to-end at least once.**
  `src/roundtrip.py` is implemented in a tiered strategy
  (`_try_direct_compile` → `_try_probe_compile` → `_make_inconclusive_issue`)
  with 16 unit tests in `tests/test_phase4_5_6.py` covering the
  Linux-reachable paths. UAT runs #8–#9 (2026-05-11) caught three
  Class-A bugs in the Windows path:
    - bare `vbproj.Compile()` raises `AttributeError: <unknown>.Compile`
      on modern Office (method hidden in type library) — fixed in
      `e1258b6`.
    - the injected `.bas` text included the export-only `Attribute
      VB_Name = "…"` header, which VBE rejects inside a module body —
      fixed in this commit via `_strip_export_directives()`.
    - `verify_compile(..., timeout_s=30)` was advertised as a timeout
      but never enforced; an invisible VBE compile-error dialog could
      hang the parent process for minutes. Fixed in this commit via
      `_run_with_timeout()` (daemon-thread worker + Office-process
      `taskkill /F` on overrun).
  `VBA_RT002` (warning) was introduced as a distinct rule_id from
  `VBA_RT000` (info, platform missing) so callers can tell whether VBE
  was reachable at all.

  **Still open:** verify on a real Office install that Strategy 2
  (probe Sub via `Application.Run`) actually forces a compile and
  surfaces real errors as `VBA_RT001` on the `tests/demo/` fixtures.
  Steps:
    1. Install Office on a Windows machine and `pip install pywin32`.
    2. Enable Office Trust Center → Macro Settings → "Trust access to
       the VBA project object model".
    3. `vbalidator tests/samples/valid_code/valid_sample.bas --host excel --roundtrip --quiet`
       — expect score=100, compile_safe=True, **zero** VBA_RT001 and
       **zero** VBA_RT002. The header strip + clean compile should
       produce an empty issue list.
    4. `vbalidator tests/demo --host excel --roundtrip --quiet`
       — expect at least one VBA_RT001 from BadModule.bas /
       BadClass.cls. If VBA_RT002 fires instead, neither Strategy 1 nor
       2 worked on this Office build and we need Strategy 3
       (VBE.CommandBars menu invocation) as a follow-up.

- [ ] **Self-hosted Windows runner for `roundtrip.yml`.**
  The workflow already exists and is wired up; it currently no-ops on
  GitHub-hosted `windows-latest` because Office is not pre-installed.
  To activate the nightly disagreement-detection alert:
    1. Provision a Windows VM with Office + pywin32 + Trust Center
       relaxed.
    2. Register it as a self-hosted runner labelled e.g.
       `windows-office`.
    3. Edit `.github/workflows/roundtrip.yml` → `runs-on: [self-hosted, windows-office]`
       (or keep `windows-latest` if your org has an Office image).
    4. The probe step (`Skip if Office is not installed on the runner`)
       will then pass and the static-vs-dynamic diff job will run.

- [ ] **Validate `tools/VBA_Model_Exporter.bas` against every supported host.**
  The macro now uses a host-agnostic resolution chain
  (`Application.VBE.ActiveVBProject` → `ThisWorkbook` → `ThisDocument` →
  `Application.CurrentProject`). It compiles in isolation but has not
  been smoke-tested in each host. Acceptance: run it once each in Excel,
  Word, Access, PowerPoint, Outlook and confirm the resulting
  `vba_references.json` contains `VBA`, the host's own library, and
  `Office`.


## B. Things that need account access (PyPI / GHCR / GitHub Pages)

All shipped during phase-5 / iter-5 — kept here as a record of the
one-time setup. If anything breaks, these are the dials to check.

- [x] **PyPI Trusted Publisher** registered (`vbalidator`) — OIDC via
  `release.yml`, `pypi` environment configured.
- [x] **GHCR** enabled and public — `ghcr.io/twobeass/vbalidator:latest`
  multi-arch (amd64 + arm64), pushed by `docker.yml`.
- [x] **GitHub Pages** deployed at <https://twobeass.github.io/VBAlidator/>
  from `docs.yml` on every `main` push.
- [x] **Branch protection on `main`** — 8 required status checks
  (`Lint (ruff)`, `Test (Py3.12 on ubuntu-latest)`, `Rule docs in sync`,
  `CLI smoke test`, `PR title is a Conventional Commit`,
  `Dependency CVE scan`, `Bandit static-security scan`,
  `CodeQL (Python)`), PRs required, force-push blocked.
  `release.yml` uses `secrets.SEMANTIC_RELEASE_TOKEN` (fine-grained
  PAT) to bypass branch protection for version-bump commits — see
  `CLAUDE.md` for the rationale.
- [x] **semantic-release tag flow** active — `v1.0.x` … `v1.4.x`
  shipped, each triggered by a `feat:` / `fix:` PR.


## C. Code-level work explicitly deferred during PR #14

These are tracked in `docs/roadmap.md` already; restating here so the
backlog is visible from the repo root.

- [x] **P2.6 — UDT / Class member-chain depth.** *(shipped 54f7919)*
  Member chains now resolve to arbitrary depth — through array indices,
  function-call returns, and property-get returns. Root cause was two
  parser bugs (`name() As T` parsed as `Variant()` because the array
  suffix was checked after `As` instead of before, in both `parse_udt`
  and `parse_declaration`). The Set/Let validator now also runs on
  dotted LHS, gated by an explicit-typing requirement. New rule:
  VBA260 (Member not found in type).

- [x] **P3.5 — Module-Level vs Procedure-Level Statement Placement.** *(shipped 29f991b)*
  Rejects `Type` / `Enum` / `Declare` / `Public Const` / `Option` /
  `DefXxx` inside a procedure body (VBA360), and rejects executable
  statements at module top level (VBA361). Parser skips the `.cls`
  `VERSION … BEGIN … END` header so class modules don't false-positive.

- [x] **VBA350 — `End Sub` closing a Function.** *(shipped 687341c)*
  `End Sub`/`End Function`/`End Property` terminator-mismatch detection
  in `procedures_parse`; explicit fixture + direct tests.

- [ ] **P5.10 — Auto-fix engine.**
  The `fix_hint` field on every `Rule` is the seed. Most useful first:
  missing `Set`, missing `PtrSafe`, wrong Property arity, Levenshtein-
  suggesting typo'd built-in names. Output: unified diff or `--fix`
  apply.

- [ ] **P5.11 — VS Code extension.**
  Separate repo. LSP server wraps `precheck()` for inline diagnostics
  on `.bas` / `.cls` / `.frm`. The `release.yml` flow can dispatch a
  `repository_dispatch` event to trigger its build.

- [ ] **P5.12 — Web demo.**
  Streamlit / Gradio app — paste code, get score. Deploys to
  Hugging Face Spaces from the same release job.


## D. Manual review / acceptance items

- [ ] **Walk through the full UAT script.** See `docs/uat.md` — covers
  install, CLI, Python API, auto-load, Score, JSON v2, MkDocs, Docker,
  every shipped rule. Sign off section by section.

- [x] **Reduce Awesome-VBA baselines toward zero.**
  Closed in iter-6 (PRs #26/#27/#28/#29/#30/#31) and iter-7 (PRs
  #33/#34/#35/#36/#37/#38/#39/#40/#41). Final baselines as of v1.4.x:
  JSONBag **0**, VBA-MemoryTools **0**, VbTrickTimer **0**, stdVBA **4**
  — and **all 4 remaining are genuine upstream library bugs** (3 typos
  in stdImage.cls + 1 missing `On Error GoTo` label in stdCallback.cls),
  not analyzer false-positives. VBAlidator's FP surface across the
  four awesome_vba projects is now **zero**. Down from the original
  203 total (-98%).

- [ ] **Performance benchmark.**
  The roadmap promises `pytest-benchmark` regression checks. Implement:
  add a tiny `tests/bench/` with a 200-line representative module,
  hook into `ci.yml` with `github-action-benchmark`, alert on ≥20%
  slowdown.

- [ ] **Confidence-score calibration set.**
  Phase 4.2 sketches a 50-module curated set with known compile
  status. Today the score formula is deterministic but never verified
  empirically. Build the set, dump score distribution, eyeball the
  >90 / <50 split.


## E. One-line follow-ups for later PRs

- [x] **Drop the `pytest>=7` floor in `pyproject.toml` to `>=8` once Py3.9
      is removed from the matrix.** *(shipped — Py3.9 EOL'd Oct 2025, fell
      out of the windows-latest setup-python cache. Matrix is now
      Py3.10–3.13, `requires-python = ">=3.10"`.)*
- [ ] Consider a `--fix` dry-run flag now that the registry has
      `fix_hint` per rule — even without auto-fix, printing the hints
      next to each issue helps humans triage.
- [ ] Move `tests/awesome_vba/` to git LFS once we add more upstream
      libraries to the regression wall.
- [ ] Convert `bug_report.md` (legacy) to a `bug.yml` structured form
      matching the style of `feature.yml` / `false_positive.yml`.

---

When in doubt: prefer a **small, scope-limited PR per item** with a
matching test or fixture — every Phase-0–5 PR was structured that way
and it made the rebase / review cycle painless.
