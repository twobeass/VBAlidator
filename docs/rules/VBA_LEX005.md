# VBA_LEX005 — Identifier too long

**Severity:** 🔴 error    **Category:** `lexer`    **Phase:** 3.9

## Description

VBA identifiers (names) are limited to 255 characters.

## Failing example

```vb
Dim aaaa…(>255 chars) As Long
```

## Compliant example

```vb
Dim total As Long
```

## How to fix

Shorten the identifier to 255 characters or fewer.

---

_Source: [src/rules.py](https://github.com/twobeass/VBAlidator/blob/main/src/rules.py) — entry `VBA_LEX005`._
