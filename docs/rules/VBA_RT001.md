# VBA_RT001 — VBE round-trip compile error

**Severity:** 🔴 compile_verified    **Category:** `roundtrip`    **Phase:** 4.5

## Description

The actual VBE compiler refused the source. This is the strongest possible verdict — a real Office host has rejected the code, so the static analyser's pass / fail call is confirmed dynamically.

## How to fix

Open the source in the VBE manually to see the full error message; the round-trip report includes the VBE description.

---

_Source: [src/rules.py](https://github.com/twobeass/VBAlidator/blob/main/src/rules.py) — entry `VBA_RT001`._
