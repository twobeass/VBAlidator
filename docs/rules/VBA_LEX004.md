# VBA_LEX004 — Unterminated string or missing end bracket

**Severity:** 🔴 error    **Category:** `lexer`    **Phase:** 3.9

## Description

A string literal has no closing `"`, or a bracket-quoted name `[…]` has no closing `]`.

## Failing example

```vb
s = "hello
```

## Compliant example

```vb
s = "hello"
```

## How to fix

Add the missing closing quote or `]`.

---

_Source: [src/rules.py](https://github.com/twobeass/VBAlidator/blob/main/src/rules.py) — entry `VBA_LEX004`._
