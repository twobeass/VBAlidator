# VBA_RT000 — Round-trip verification unavailable

**Severity:** 🔵 info    **Category:** `roundtrip`    **Phase:** 4.5

## Description

The runtime could not even attempt a VBE round-trip — usually because we're not on Windows, pywin32 is missing, or Office is not installed. Static analysis remains the authoritative result; this is informational only.

## How to fix

Install Microsoft Office and `pip install pywin32` to enable round-trip verification, or simply drop `--roundtrip` from the invocation.

---

_Source: [src/rules.py](https://github.com/twobeass/VBAlidator/blob/main/src/rules.py) — entry `VBA_RT000`._
