# VBA_LEX001 — Unexpected character

**Severity:** 🔴 error    **Category:** `lexer`    **Phase:** 0

## Description

The lexer encountered a character that is not part of any valid VBA token. Often an encoding mistake (smart quotes, Euro sign, …) introduced by copy-paste.

## Failing example

```vb
x = 1€
```

## Compliant example

```vb
x = 1
```

## How to fix

Replace the offending character with its ASCII equivalent or remove it.

---

_Source: [src/rules.py](../../src/rules.py) — entry `VBA_LEX001`._
