# VBAlidator rule catalogue

Stable rule IDs emitted by VBAlidator. Each row links to the rule's detail page. Use the `rule_id` to silence specific findings in your CI ignore list — the IDs do not change between releases.

| Rule | Severity | Category | Phase | Title |
| --- | --- | --- | --- | --- |
| [`VBA001`](VBA001.md) | 🔴 error | `name_resolution` | 0 | Undefined identifier |
| [`VBA002`](VBA002.md) | 🔴 error | `member_access` | 0 | Member not found |
| [`VBA003`](VBA003.md) | 🔴 error | `declaration` | 0 | Duplicate declaration |
| [`VBA004`](VBA004.md) | 🔴 error | `member_access` | 0 | Invalid `.member` reference without With |
| [`VBA005`](VBA005.md) | 🔴 error | `type` | 0 | Expected Array or Procedure |
| [`VBA006`](VBA006.md) | 🔴 error | `signature` | 0 | Argument count mismatch |
| [`VBA007`](VBA007.md) | 🔴 error | `signature` | 0 | ByRef argument type mismatch |
| [`VBA008`](VBA008.md) | 🔴 error | `control_flow` | 0 | Exit statement type mismatch |
| [`VBA009`](VBA009.md) | 🟡 warning | `control_flow` | 0 | Unreachable code |
| [`VBA010`](VBA010.md) | 🔴 error | `syntax` | 0 | Syntax error |
| [`VBA101`](VBA101.md) | 🔴 error | `declaration` | 1.4 | ReDim target undefined |
| [`VBA102`](VBA102.md) | 🔴 error | `declaration` | 1.4 | ReDim target not a variable |
| [`VBA103`](VBA103.md) | 🔴 error | `declaration` | 1.4 | ReDim target is not a dynamic array |
| [`VBA104`](VBA104.md) | 🔴 error | `declaration` | 1.4 | Erase target undefined |
| [`VBA105`](VBA105.md) | 🔴 error | `declaration` | 1.4 | Erase target not a variable |
| [`VBA106`](VBA106.md) | 🔴 error | `declaration` | 1.4 | Erase target is not an array |
| [`VBA201`](VBA201.md) | 🔴 error | `jump` | 2.1 | Jump target is not a label |
| [`VBA210`](VBA210.md) | 🔴 error | `assignment` | 2.2 | `Set` used on a scalar target |
| [`VBA211`](VBA211.md) | 🔴 error | `assignment` | 2.2 | Object assignment without `Set` |
| [`VBA221`](VBA221.md) | 🔴 error | `property` | 2.3 | Property Let/Set has zero parameters |
| [`VBA222`](VBA222.md) | 🔴 error | `property` | 2.3 | Property Let/Set arity disagrees with Get |
| [`VBA223`](VBA223.md) | 🔴 error | `property` | 2.3 | Property Set value parameter is scalar (use Let) |
| [`VBA224`](VBA224.md) | 🔴 error | `property` | 2.3 | Property Let value parameter is object (use Set) |
| [`VBA230`](VBA230.md) | 🔴 error | `const_expression` | 2.5 | `Const` initialiser calls a function |
| [`VBA231`](VBA231.md) | 🔴 error | `const_expression` | 2.5 | `Const` initialiser references a non-constant |
| [`VBA240`](VBA240.md) | 🔴 error | `operator_type` | 2.4 | Arithmetic operator between string and numeric literal |
| [`VBA250`](VBA250.md) | 🔴 error | `declaration` | 2.8 | Fixed-length String at procedure level |
| [`VBA300`](VBA300.md) | 🔴 error | `platform` | 3.3 | `Declare` missing `PtrSafe` |
| [`VBA310`](VBA310.md) | 🔴 error | `enum` | 3.4 | Duplicate Enum member name |
| [`VBA320`](VBA320.md) | 🟡 warning | `style` | 3.6 | Module is missing `Option Explicit` |
| [`VBA330`](VBA330.md) | 🔴 error | `interface` | 3.1 | Class is missing an interface method |
| [`VBA340`](VBA340.md) | 🔴 error | `events` | 3.2 | `RaiseEvent` target has no matching `Event` declaration |
| [`VBA341`](VBA341.md) | 🔴 error | `events` | 3.2 | `RaiseEvent` argument count mismatch |
| [`VBA350`](VBA350.md) | 🔴 error | `syntax` | 3.5 | Procedure terminator does not match its kind |
| [`VBA_LEX001`](VBA_LEX001.md) | 🔴 error | `lexer` | 0 | Unexpected character |
| [`VBA_LEX002`](VBA_LEX002.md) | 🔴 error | `lexer` | 2.7 | Invalid date literal |
| [`VBA_RT000`](VBA_RT000.md) | 🔵 info | `roundtrip` | 4.5 | Round-trip verification unavailable |
| [`VBA_RT001`](VBA_RT001.md) | 🔴 compile_verified | `roundtrip` | 4.5 | VBE round-trip compile error |
| [`VBA_RT002`](VBA_RT002.md) | 🟡 warning | `roundtrip` | 4.5 | Round-trip verification inconclusive |

*39 rules registered.* Generated from `src/rules.py` via `python tools/generate_rule_docs.py`.
