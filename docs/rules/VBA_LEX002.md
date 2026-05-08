# VBA_LEX002 — Invalid date literal

**Severity:** 🔴 error    **Category:** `lexer`    **Phase:** 2.7

## Description

A `#…#` literal does not parse as any of the recognised VBA date / time formats, or its month / day / hour fields are out of range.

## Failing example

```vb
d = #2025-13-45#
```

## Compliant example

```vb
d = #2025-01-15#
```

## How to fix

Use m/d/y, yyyy-mm-dd, d-mmm-y, or 'MMMM d, y' format with valid date components.

---

_Source: [src/rules.py](../../src/rules.py) — entry `VBA_LEX002`._
