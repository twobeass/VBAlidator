# VBA_RT002 — Round-trip verification inconclusive

**Severity:** 🟡 warning    **Category:** `roundtrip`    **Phase:** 4.5

## Description

The runtime tried to drive the VBE compiler but no trigger succeeded — `VBProject.Compile()` is hidden on modern Office, and the probe-Sub via `Application.Run` either timed out or failed with an unrecognised description. Distinct from `VBA_RT000`: VBE *was* reachable, we just couldn't reach a verdict.

## How to fix

See TODO.md §A2 for the open work on Strategy 3 (VBE menu-bar invocation). In the meantime: rely on the static analyser, which remains the authoritative answer.

---

_Source: [src/rules.py](https://github.com/twobeass/VBAlidator/blob/main/src/rules.py) — entry `VBA_RT002`._
