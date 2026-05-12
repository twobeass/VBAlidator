# Host-model false-positive comparison

The bundled `src/models/{excel,word,access,outlook}.json` are hand-
curated 80-percent subsets. A separate UAT run on Windows + Office
365 used `tools/generate_model.py` to introspect the real
type libraries and emit full-fidelity counterparts (~250× larger).
This page captures the head-to-head FP comparison and the design
decision it informs.

Per-project error counts from `precheck(<project>, model_path=…)`
with `severity=error` only, run on each Awesome-VBA project:

| Project | Files | Baseline (`std_model` only) | Shipped Excel | Regen Excel |
|---|---:|---:|---:|---:|
| JSONBag | 2 | 6 | 12 | **6** |
| VBA-MemoryTools | 4 | 13 | 28 | 28 |
| VbTrickTimer | 1 | 5 | 5 | 5 |
| stdVBA | 25 | 224 | 306 | **267** |

Cross-host totals (`min(host_total)` for each side, lower = fewer FPs):

|     | excel | word | access | visio |
|---|---:|---:|---:|---:|
| **Shipped** | **351** | 357 | 377 | — |
| **Regenerated** | **306** | 322 | 348 | 332 |

The regenerated model is **strictly better or equal** on every
project — no regression introduced anywhere. Best-Excel improves by
**45 errors (−12.8%)** mostly from `stdVBA-master` (`306 → 267`) and
JSONBag (`12 → 6`).

Notable: the shipped models *worsen* JSONBag versus the
`std_model.json`-only baseline (6 → 12). The Workbook / Worksheet
class definitions in the shipped curated model register members that
shadow JSONBag's project-internal `Names` / `Item` references; the
regenerated model exposes the same members but also the full
sibling set, so the analyser's strict-shadowing heuristic resolves
both ways correctly.

## Decision

**Plan C — opt-in full models** is viable. Concretely:

1. Ship a separate sdist `vbalidator-models-full` (~7 MB) that drops
   the full models into a `vbalidator.models.full` namespace
   package. Users opt in via `pip install vbalidator[full-models]`
   (the extra resolves to the standalone package via setuptools'
   `optional-dependencies`).

2. The CLI gains a `*-full` variant of `--host`:
   `--host excel-full` first probes
   `importlib.resources.files("vbalidator.models.full") / "excel.json"`
   and falls back to the bundled curated model when that resource is
   absent. No behavioural change for users on the lean install.

3. Documentation: a new `docs/full-models.md` page (linked from
   `docs/Configuration.md`) explains the trade-off — coverage vs.
   install size — and points to this comparison.

The work is **deferred to a follow-up PR** to keep
`claude/improve-vba-precompiler-3YcZp` scoped to roundtrip-pipeline
fixes. Today's contract:

- The shipped curated 80-percent subsets remain the default.
- `tools/generate_model.py` lets any user generate full-fidelity
  models themselves from a Windows + Office host (see
  `docs/Configuration.md` — Generating a model from COM).
- The FP gate is recorded here so the design discussion has data,
  not speculation.

## Reproducing the comparison

The gist of regenerated artefacts lives at
<https://gist.github.com/twobeass/6786ef3404922c3549d5621638be29e6>.
Clone it locally and run:

```bash
git clone https://gist.github.com/twobeass/6786ef3404922c3549d5621638be29e6.git /tmp/gist_clone
# Then, from the VBAlidator repo root:
python3 - <<'PY'
import sys
from pathlib import Path
sys.path.insert(0, ".")
from src.api import precheck

HOSTS = ("excel", "word", "access", "visio")
SHIPPED = Path("src/models")
REGEN = Path("/tmp/gist_clone")
for project in sorted(p for p in Path("tests/awesome_vba").iterdir() if p.is_dir()):
    base = precheck(project)
    print(f"{project.name:25s} baseline={len(base.errors)}")
    for host in HOSTS:
        ship = SHIPPED / f"{host}.json"
        regen = REGEN / f"{host}.json"
        if ship.is_file():
            r = precheck(project, model_path=str(ship))
            print(f"    shipped/{host:<8s} errors={len(r.errors)}")
        if regen.is_file():
            r = precheck(project, model_path=str(regen))
            print(f"    regen/{host:<8s}   errors={len(r.errors)}")
PY
```

Re-run after any change that touches `src/models/`, `src/analyzer.py`
or the host-model schema — a future commit that increases FPs on the
regenerated side without improving the shipped side is a regression
worth investigating.
