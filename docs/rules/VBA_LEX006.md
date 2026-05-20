# VBA_LEX006 — Line too long

**Severity:** 🔴 error    **Category:** `lexer`    **Phase:** 3.9

## Description

A physical source line exceeds VBA's 1023-character limit.

## Failing example

```vb
' a single line longer than 1023 characters …
```

## Compliant example

```vb
' wrap with the line-continuation character `_`
```

## How to fix

Break the line using the `_` line-continuation character.

---

_Source: [src/rules.py](https://github.com/twobeass/VBAlidator/blob/main/src/rules.py) — entry `VBA_LEX006`._
