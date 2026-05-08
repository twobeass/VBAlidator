# TODO — what still needs human / Windows / live-account access

The Phase 0–5 work in PR #14 is complete from the analyser's side, but a
handful of items can only be done with **a real Windows + Office host**,
**live PyPI / GHCR / GitHub-Pages accounts**, or **deliberate human
review**. They are listed here so they don't get lost, grouped by who can
do them.

Status legend: `[ ]` open · `[~]` partially done in the sandbox · `[x]` done.


## A. Things that need a Windows + Office machine

These cannot run on the Linux GitHub-hosted runners.

- [ ] **Regenerate every bundled host model from a real Office install.**
  The `src/models/{excel,word,access,outlook}.json` files shipped today
  are hand-curated minimal subsets covering the ~80% common case
  (Application / Workbook / Worksheet / Range / Documents / Database /
  NameSpace / MailItem). For full fidelity:
    1. Open Excel → VBE → File ▶ Import File → `tools/VBA_Model_Exporter.bas`.
    2. Run `ExportReferences`. It writes `vba_references.json` next to
       the workbook (or to `%TEMP%`).
    3. On the same machine: `pip install comtypes` → `python tools/generate_model.py vba_references.json -o src/models/excel.json`.
    4. Repeat for Word, Access, Outlook (the script writes to the path
       you give it).
    5. Diff against the shipped models, verify, commit.

- [ ] **Run the VBE round-trip suite end-to-end at least once.**
  `src/roundtrip.py` is implemented and unit-tested (off-platform fallback
  path), but nobody has yet driven it against a real VBE compiler. Steps:
    1. Install Office on a Windows machine and `pip install pywin32`.
    2. Enable Office Trust Center → Macro Settings → "Trust access to the
       VBA project object model".
    3. From a clone of this branch run
       `vbalidator tests/samples/valid_code/valid_sample.bas --host excel --roundtrip --quiet`
       and confirm the result is `score=100, compile_safe=True` with no
       `VBA_RT001` issues (only the `VBA_RT000` info is acceptable, and
       only if Trust Center is locked).
    4. Repeat with `tests/demo` — the demo files contain known errors
       and VBE itself should reject several modules; verify VBA_RT001
       fires for at least the `BadModule.bas` / `BadClass.cls` ones.

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

- [ ] **P2.6 — UDT / Class member-chain depth.**
  Today member chains validate the first hop fully and degrade to
  permissive Variant past the second `.`. Closing this requires a
  proper type-system on the analyser walker. Estimated: 1 PR, +400/-50
  in `src/analyzer.py` + ~10 fixtures.

- [ ] **P3.5 — Module-Level vs Procedure-Level Statement Placement.**
  Reject `Type … End Type` inside a `Sub`, executable code at module
  level, etc. Needs the parser to track the active container. Probably
  ~250 LOC, 4–6 fixtures.

- [ ] **VBA350 — `End Sub` closing a Function.**
  Easy rule, AI generators sometimes get this wrong. ~30 LOC + 1
  fixture.

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

- [ ] Drop the `pytest>=7` floor in `pyproject.toml` to `>=8` once Py3.9
      is removed from the matrix.
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
