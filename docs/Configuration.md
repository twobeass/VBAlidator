# Configuration

VBAlidator reads its symbol knowledge from three layered sources, in
the following precedence (highest first):

1. **Custom model** — passed via `--model my.json` or
   `precheck(model_path="my.json")`.
2. **Host model** — bundled with the package, selected by
   `--host excel|word|access|outlook` or
   `precheck(host="excel")`. Lives at `src/models/<host>.json`.
3. **Standard model** — `src/std_model.json`. Always loaded; covers
   the VBA runtime (string functions, math, dates, file IO, vb*
   constants).

Anything resolved at level 1 wins over level 2 and level 3. Identifiers
not resolved at any layer trigger `VBA001 Undefined identifier`.

## Conditional-compilation defaults

`Config()` ships with modern Microsoft 365 defaults:

| Constant | Default | Override |
|----------|---------|----------|
| `VBA7` | True | `--define VBA7=False` |
| `WIN64` | True | `--define WIN64=False` |
| `WIN32` | False | `--define WIN32=True` |
| `WIN16` | False | (legacy only) |
| `MAC` | False | `--define MAC=True` |

VBA's `#If` evaluator is **case-insensitive**, so
`#If Vba7 Then` and `#If VBA7 Then` are equivalent.

## Bundled host models

| `--host` | File | Highlights |
|----------|------|------------|
| `excel` | `src/models/excel.json` | Application, Workbook, Worksheet, Range, WorksheetFunction, Names, Shapes, Window, Interior, Font, Borders, Validation, Chart + ~70 `xl*` enum aliases |
| `word` | `src/models/word.json` | Application, Documents, Document, Range, Selection, Window + `wd*` aliases |
| `access` | `src/models/access.json` | Application, DoCmd, Database, Recordset, DBEngine, CurrentProject, DLookup/DCount/DSum/Nz globals + `ac*` / `db*` aliases |
| `outlook` | `src/models/outlook.json` | Application, NameSpace, Folder, Items, MailItem + `ol*` aliases |

## Custom models

A custom JSON model lets you cover libraries / classes the bundled
host models don't ship — e.g. Visio, AutoCAD, ComctlLib, your own
add-in's `.tlb`.

### Schema

```jsonc
{
  "globals": {
    "MyGlobal": {
      "type": "Function",          // or a class name
      "returns": "Variant",        // when type=="Function"
      "min_args": 1,
      "max_args": 3,
      "args": [                    // optional; required for ByRef checks
        { "name": "x", "type": "Long", "mechanism": "ByRef" }
      ]
    }
  },
  "classes": {
    "MyClass": {
      "members": {
        "DoStuff": { "type": "Sub" },
        "Value":   { "type": "Long" }
      }
    }
  },
  "enums": {
    "MyEnum": { "Foo": 0, "Bar": 1 }
  },
  "references": [
    { "name": "Visio" }
  ]
}
```

`globals` keys land in the global scope; `classes` describe member
chains for `--host` types; `enums` register both the enum name and
each member as a globally visible Long; `references` register library
names so qualified accesses like `Visio.Application` resolve.

### Generating a model from COM

Two-step workflow:

1. **In your host VBE** — import `tools/VBA_Model_Exporter.bas` into
   any Office host (Excel, Word, Access, PowerPoint, Outlook, Visio,
   AutoCAD, …) and run `ExportReferences`. It walks
   `Application.VBE.ActiveVBProject.References` and writes a
   `vba_references.json` next to the open document. Requires
   "Trust access to the VBA project object model" in the Trust Center.

2. **On your dev machine (Windows + comtypes)**:

   ```bash
   pip install comtypes
   python tools/generate_model.py path/to/vba_references.json -o vba_model.json
   ```

   The script introspects every type library, captures classes / enums /
   constants, copies CoClass default-interface members, and lifts the
   first-found Application interface into the global scope. Use
   `--no-app-promote` to skip the global lift.

3. **Pass it to VBAlidator** — either explicitly:

   ```bash
   vbalidator ./MyAddin --host excel --model ./vba_model.json
   ```

   …or implicitly: drop the file as `vba_model.json` next to the input
   folder (or in your CWD) and VBAlidator picks it up automatically.

   ```bash
   vbalidator ./MyAddin --host excel    # auto-loads ./vba_model.json
   ```

`std_model.json` and `--host` always merge first; `--model` (or the
auto-detected `vba_model.json`) is layered on top and may override
individual entries.

A reference Visio export ships at
[`examples/vba_references.example.json`](https://github.com/twobeass/VBAlidator/blob/main/examples/vba_references.example.json) — feed it through
`generate_model.py` to see what a custom model looks like end-to-end.

## Built-in heuristics

### 1. Form-control dynamic resolution
In `.frm` modules, undefined identifiers default to `Object` so
implicit form controls (`Me.lblStatus.Caption = …`) don't trip the
analyser. Members that *are* in the standard `UserForm` model still
resolve normally.

### 2. Strict project shadowing
A project module named `Excel` shadows the `Excel` library. When a
type is a known project module, member lookup is strict against that
module — fallback to libraries does not happen, so a typo'd member
inside your own module still raises `VBA002 Member not found`.

### 3. UserForm / ThisDocument fallback
Members not found on a Form fall back to base `UserForm`. `ThisDocument`
falls back to `Document` / `IVDocument`.

### 4. Default member resolution
`obj(1)` is implicitly resolved to `obj.Item(1)` if the type exposes
`Item`. Lets `Selection(1)` correctly type-resolve to `Shape` etc.

### 5. Identifier suffix normalisation
`Mid$`, `Mid`, and `[Mid]` resolve to the same standard global. The
suffix-stripping is purely a lookup convenience — the *original*
spelling is preserved in error messages.

### 6. Case-insensitive lookup
All model lookups are normalised to lowercase, matching VBA semantics.

## Reserved keywords
The following identifiers are reserved by the analyser. Using them as
variable names triggers parser-level errors via the legacy `VBA010`
rule.

| Category | Keywords |
|----------|----------|
| Flow | `if`, `then`, `else`, `elseif`, `end`, `exit`, `for`, `next`, `do`, `loop`, `while`, `wend`, `select`, `case`, `with`, `goto`, `gosub`, `resume`, `stop`, `on`, `error` |
| Declarations | `dim`, `static`, `const`, `public`, `private`, `global`, `friend`, `sub`, `function`, `property`, `get`, `let`, `set`, `type`, `new`, `withevents`, `as`, `implements`, `event`, `raiseevent` |
| Types | `boolean`, `integer`, `long`, `longlong`, `longptr`, `single`, `double`, `currency`, `decimal`, `date`, `string`, `variant`, `object`, `byte`, `nothing`, `empty`, `null` |
| Operators | `and`, `or`, `xor`, `not`, `is`, `like`, `typeof`, `mod`, `true`, `false`, `eqv`, `imp` |
| Arrays / IO | `redim`, `preserve`, `erase`, `open`, `close`, `input`, `output`, `append`, `binary`, `random`, `put`, `print` |
| Misc | `len`, `mid`, `call`, `defint`, `defstr`, `defbool`, `deflng`, `defdbl`, `defsng`, `defcur`, `defdate`, `defobj`, `defvar`, `option` |

## Implicit top-level globals
- Forms / Classes with `Attribute VB_PredeclaredId = True` are
  registered as a global instance named after the class.
- The current host (Excel `Application`, Word `Application`, …) lives
  in the matching `--host` model and so is always available when the
  flag is set.
