# Host-model coverage

Since v1.1.0 the four Office hosts (Excel / Word / Access / Visio)
ship as **full-fidelity** type-library extracts rather than the
hand-curated 80-percent subsets that earlier 0.x / 1.0.x versions
used. The earlier "Plan C — opt-in `*-full` package" comparison this
page used to host is no longer relevant; the full models are the
default.

If you want to verify the host-model FP delta on your own corpus, the
shape is:

```python
from src.api import precheck

# Baseline — std_model only
no_host = precheck("path/to/project")

# With the Office host layered in
with_excel = precheck("path/to/project", host="excel")

print(f"std-only: {len(no_host.errors)} errors")
print(f"excel:    {len(with_excel.errors)} errors")
```

For the bundled `awesome_vba` regression fixtures (JSONBag, stdVBA,
VBA-MemoryTools, VbTrickTimer) the iter-6 + iter-7 cycles drove the
total from **203 → 4 hard errors** (-98 %), with all four remaining
findings being genuine upstream library bugs rather than
analyzer false-positives. See
[`tests/test_awesome_vba_regression.py::BASELINE`](https://github.com/twobeass/VBAlidator/blob/main/tests/test_awesome_vba_regression.py)
for the per-project reasons.

## Companion COM stubs (auto-layering)

In addition to the five Office hosts, six companion stubs
(`mscomctl`, `msforms`, `scripting`, `vbscript_regexp`,
`wscript_shell`, `shell_application`) load automatically when the
scan set mentions the matching ProgID / namespace. See
[Configuration → Bundled host models](Configuration.md#bundled-host-models)
for the full trigger table.

## Regenerating the Office models

```bash
# Inside Office: VBE → File ▶ Import File → tools/VBA_Model_Exporter.bas
# Run ExportReferences → writes vba_references.json next to the workbook.

pip install comtypes  # Windows + the target host installed
python tools/generate_model.py vba_references.json -o vba_model.json
```

The four scripting/shell COM stubs (`scripting.json`,
`vbscript_regexp.json`, `wscript_shell.json`, `shell_application.json`)
are hand-curated — small enough that maintaining them by hand stays
easier than driving comtypes against their type libraries.
The MSComCtl / MSForms stubs are comtypes extracts plus a VB6-
container-control-member patch; rebuild via
`tools/build_mscomctl_model.py` / `tools/build_msforms_model.py`.
