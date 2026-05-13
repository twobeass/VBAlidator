# TODO — what still needs human / Windows / live-account access

The Phase 0–5 work in PR #14 is complete from the analyser's side, but a
handful of items can only be done with **a real Windows + Office host**,
**live PyPI / GHCR / GitHub-Pages accounts**, or **deliberate human
review**. They are listed here so they don't get lost, grouped by who can
do them.

Status legend: `[ ]` open · `[~]` partially done in the sandbox · `[x]` done.


## A. Things that need a Windows + Office machine

These cannot run on the Linux GitHub-hosted runners.

- [~] **Regenerate every bundled host model from a real Office install.**
  The `src/models/{excel,word,access,outlook}.json` files shipped today
  are hand-curated minimal subsets covering the ~80% common case
  (Application / Workbook / Worksheet / Range / Documents / Database /
  NameSpace / MailItem). UAT run #9 produced full-fidelity counterparts
  on Windows + Office 365 (gist at
  <https://gist.github.com/twobeass/6786ef3404922c3549d5621638be29e6>);
  the FP comparison documented in `docs/host-models-comparison.md`
  shows the regenerated models are **strictly better or equal** on every
  Awesome-VBA project (best-Excel total 306 vs. shipped 351 = −12.8%).

  **Decision (Plan C, opt-in full models):**
    1. Ship a separate `vbalidator-models-full` sdist (~7 MB) that
       drops the regenerated JSONs into `vbalidator.models.full`.
    2. CLI gains `--host {excel,word,access,outlook,visio}-full`;
       the standard hosts stay bundled and remain the default.
    3. Documentation in `docs/full-models.md` (linked from
       Configuration) explains the trade-off.

  **Deferred to a follow-up PR** to keep the current branch focused on
  round-trip pipeline fixes; the FP gate result is captured so the
  next person picks up with data, not speculation.

  For the **interim manual workflow** (still useful as the regen step
  for the future opt-in package):
    1. Open Excel → VBE → File ▶ Import File → `tools/VBA_Model_Exporter.bas`.
    2. Run `ExportReferences`. It writes `vba_references.json` next to
       the workbook (or to `%TEMP%`).
    3. On the same machine: `pip install comtypes` → `python tools/generate_model.py vba_references.json -o vba_model.json`.
    4. `vbalidator … --model vba_model.json` (or drop the file next to
       your code for auto-load).

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

These are one-time setup steps. After they are in place CI takes over.

- [ ] **Register the project on PyPI as a Trusted Publisher.**
  `release.yml` already uses OIDC and refuses to use API tokens.
  Setup:
    1. Reserve the project name on PyPI: <https://pypi.org/manage/projects/>
       → "Add a new project" → `vbalidator`.
    2. → Settings → Publishing → "Add a new pending publisher".
       - Owner: `twobeass`
       - Repository: `VBAlidator`
       - Workflow filename: `release.yml`
       - Environment name: `pypi`
    3. Same again on <https://test.pypi.org> for the dev-release flow,
       environment name `testpypi` (the workflow currently doesn't
       declare it; if you want OIDC for TestPyPI too, add
       `environment: testpypi` to the `publish-testpypi-nightly` job).
    4. In the GitHub repo → Settings → Environments → "New environment"
       called `pypi` (matching the workflow's `environment.name`).

- [ ] **Enable GHCR for the package.**
  `docker.yml` pushes to `ghcr.io/twobeass/vbalidator`. After the first
  successful push the package is created; you need to:
    1. Visit <https://github.com/users/twobeass/packages/container/vbalidator>.
    2. → Package settings → Manage Actions access → add this repository
       with `Write` permission.
    3. Optionally → Change visibility to Public so users can `docker pull`
       without auth.

- [ ] **Enable GitHub Pages.**
  `docs.yml` is wired up to deploy the MkDocs Material site. Activate it:
    1. Repo → Settings → Pages → Source: "GitHub Actions".
    2. Push any commit to `main` → the `Deploy to GitHub Pages` job
       publishes <https://twobeass.github.io/VBAlidator/>.

- [ ] **Configure branch protection on `main`.**
  Recommended in `docs/ci-cd.md`. Concrete steps:
    1. Repo → Settings → Branches → Add rule for `main`.
    2. ☑ Require a pull request before merging (1 reviewer).
    3. ☑ Require status checks to pass:
       - `Lint (ruff)`
       - `Test (Py3.12 on ubuntu-latest)` (and any other matrix entry
         you want to gate on)
       - `Rule docs in sync`
       - `CLI smoke test`
       - `PR title is a Conventional Commit`
       - `pip-audit`
       - `Bandit static-security scan`
       - `CodeQL (Python)`
    4. ☑ Require signed commits.
    5. ☑ Do not allow force-push.

- [ ] **Trigger the first semantic-release tag.**
  Until at least one commit lands on `main` with a `feat:` / `fix:` /
  `BREAKING CHANGE:` footer, `release.yml` will not produce a tag.
  After the PR is merged, push a single commit (e.g.
  `chore(release): seed v0.2.0`) or open a tiny `feat:` PR; the tag
  flow then runs end-to-end and the version becomes the source of
  truth for `pyproject.toml` + `src/__init__.__version__`.


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

- [ ] **Reduce Awesome-VBA baselines toward zero.**
  Today's ceilings (`tests/test_awesome_vba_regression.py`):
  JSONBag 12, VBA-MemoryTools 18, VbTrickTimer 5, stdVBA 335. Each
  remaining error is either:
    - A real bug in upstream → file a PR there.
    - A genuine VBAlidator gap → register a `# noqa`-style suppression
      hint or close the analyser hole.

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
