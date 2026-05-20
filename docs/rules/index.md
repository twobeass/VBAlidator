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
| [`VBA110`](VBA110.md) | 🔴 error | `syntax` | 1.5 | For without Next |
| [`VBA111`](VBA111.md) | 🔴 error | `syntax` | 1.5 | Do without Loop |
| [`VBA112`](VBA112.md) | 🔴 error | `syntax` | 1.5 | While without Wend |
| [`VBA113`](VBA113.md) | 🔴 error | `syntax` | 1.5 | Expected End With |
| [`VBA114`](VBA114.md) | 🔴 error | `syntax` | 1.5 | Select Case without End Select |
| [`VBA115`](VBA115.md) | 🔴 error | `syntax` | 1.5 | Block If without End If |
| [`VBA116`](VBA116.md) | 🔴 error | `syntax` | 1.5 | Next without For |
| [`VBA117`](VBA117.md) | 🔴 error | `syntax` | 1.5 | Loop without Do |
| [`VBA118`](VBA118.md) | 🔴 error | `syntax` | 1.5 | Wend without While |
| [`VBA119`](VBA119.md) | 🔴 error | `syntax` | 1.5 | Block terminator without matching opener |
| [`VBA120`](VBA120.md) | 🟡 warning | `type` | 1.6 | User-defined type not defined |
| [`VBA122`](VBA122.md) | 🔴 error | `type` | 1.6 | User-defined type without members |
| [`VBA123`](VBA123.md) | 🔴 error | `type` | 1.6 | Empty Enum not allowed |
| [`VBA130`](VBA130.md) | 🔴 error | `signature` | 2.10 | Required parameter after Optional |
| [`VBA131`](VBA131.md) | 🔴 error | `signature` | 2.10 | ParamArray is not the last parameter |
| [`VBA132`](VBA132.md) | 🔴 error | `signature` | 2.10 | More than one ParamArray |
| [`VBA133`](VBA133.md) | 🔴 error | `signature` | 2.10 | ParamArray combined with Optional |
| [`VBA134`](VBA134.md) | 🔴 error | `signature` | 2.10 | ParamArray element type is not Variant |
| [`VBA136`](VBA136.md) | 🔴 error | `signature` | 2.10 | Array parameter passed ByVal |
| [`VBA137`](VBA137.md) | 🔴 error | `signature` | 2.10 | User-defined type passed ByVal |
| [`VBA140`](VBA140.md) | 🔴 error | `signature` | 2.9 | Named argument not found |
| [`VBA141`](VBA141.md) | 🔴 error | `signature` | 2.9 | Named argument specified more than once |
| [`VBA142`](VBA142.md) | 🔴 error | `signature` | 2.9 | Named arguments with ParamArray |
| [`VBA150`](VBA150.md) | 🔴 error | `declaration` | 2.12 | Ambiguous name detected (duplicate procedure) |
| [`VBA151`](VBA151.md) | 🔴 error | `syntax` | 2.12 | Code after End Sub/Function/Property |
| [`VBA152`](VBA152.md) | 🔴 error | `property` | 2.12 | Property Get/Let type inconsistent |
| [`VBA153`](VBA153.md) | 🔴 error | `interface` | 2.12 | Implements method signature mismatch |
| [`VBA160`](VBA160.md) | 🔴 error | `assignment` | 2.13 | Assignment to constant not permitted |
| [`VBA164`](VBA164.md) | 🔴 error | `assignment` | 2.13 | Invalid use of Me keyword |
| [`VBA165`](VBA165.md) | 🔴 error | `declaration` | 2.13 | Invalid use of New keyword |
| [`VBA170`](VBA170.md) | 🔴 error | `control_flow` | 2.11 | Exit For not within For...Next |
| [`VBA171`](VBA171.md) | 🔴 error | `control_flow` | 2.11 | Exit Do not within Do...Loop |
| [`VBA172`](VBA172.md) | 🔴 error | `control_flow` | 2.11 | Next control variable mismatch |
| [`VBA173`](VBA173.md) | 🔴 error | `control_flow` | 2.11 | For control variable already in use |
| [`VBA174`](VBA174.md) | 🔴 error | `control_flow` | 2.11 | For Each control variable must be Variant or object |
| [`VBA175`](VBA175.md) | 🔴 error | `control_flow` | 2.11 | Duplicate label |
| [`VBA180`](VBA180.md) | 🔴 error | `placement` | 3.7 | Statement invalid inside Enum block |
| [`VBA181`](VBA181.md) | 🔴 error | `placement` | 3.7 | Statement invalid inside Type block |
| [`VBA182`](VBA182.md) | 🔴 error | `placement` | 3.7 | Deftype statement must precede declarations |
| [`VBA183`](VBA183.md) | 🔴 error | `placement` | 3.7 | Duplicate Deftype letter range |
| [`VBA184`](VBA184.md) | 🔴 error | `placement` | 3.7 | Duplicate Option statement |
| [`VBA190`](VBA190.md) | 🔴 error | `preprocessor` | 3.8 | Conditional directive without matching #If |
| [`VBA191`](VBA191.md) | 🔴 error | `preprocessor` | 3.8 | Missing #End If |
| [`VBA192`](VBA192.md) | 🔴 error | `preprocessor` | 3.8 | Invalid #Const declaration |
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
| [`VBA270`](VBA270.md) | 🔴 error | `declaration` | 2.14 | Type-declaration character does not match declared type |
| [`VBA300`](VBA300.md) | 🔴 error | `platform` | 3.3 | `Declare` missing `PtrSafe` |
| [`VBA310`](VBA310.md) | 🔴 error | `enum` | 3.4 | Duplicate Enum member name |
| [`VBA320`](VBA320.md) | 🟡 warning | `style` | 3.6 | Module is missing `Option Explicit` |
| [`VBA330`](VBA330.md) | 🔴 error | `interface` | 3.1 | Class is missing an interface method |
| [`VBA340`](VBA340.md) | 🔴 error | `events` | 3.2 | `RaiseEvent` target has no matching `Event` declaration |
| [`VBA341`](VBA341.md) | 🔴 error | `events` | 3.2 | `RaiseEvent` argument count mismatch |
| [`VBA350`](VBA350.md) | 🔴 error | `syntax` | 3.5 | Procedure terminator does not match its kind |
| [`VBA360`](VBA360.md) | 🔴 error | `placement` | 3.5 | Module-only declaration inside a procedure body |
| [`VBA361`](VBA361.md) | 🔴 error | `placement` | 3.5 | Executable statement at module level |
| [`VBA370`](VBA370.md) | 🔴 error | `placement` | 3.10 | Constant/array/Declare as Public member of object module |
| [`VBA_LEX001`](VBA_LEX001.md) | 🔴 error | `lexer` | 0 | Unexpected character |
| [`VBA_LEX002`](VBA_LEX002.md) | 🔴 error | `lexer` | 2.7 | Invalid date literal |
| [`VBA_LEX004`](VBA_LEX004.md) | 🔴 error | `lexer` | 3.9 | Unterminated string or missing end bracket |
| [`VBA_LEX005`](VBA_LEX005.md) | 🔴 error | `lexer` | 3.9 | Identifier too long |
| [`VBA_LEX006`](VBA_LEX006.md) | 🔴 error | `lexer` | 3.9 | Line too long |
| [`VBA_RT000`](VBA_RT000.md) | 🔵 info | `roundtrip` | 4.5 | Round-trip verification unavailable |
| [`VBA_RT001`](VBA_RT001.md) | 🔴 compile_verified | `roundtrip` | 4.5 | VBE round-trip compile error |
| [`VBA_RT002`](VBA_RT002.md) | 🟡 warning | `roundtrip` | 4.5 | Round-trip verification inconclusive |

*90 rules registered.* Generated from `src/rules.py` via `python tools/generate_rule_docs.py`.
