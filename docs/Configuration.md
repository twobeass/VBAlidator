# Configuration & Extending VBAlidator

VBAlidator supports dynamic loading of object models (Classes, Interfaces, Enums) via external JSON. This allows you to validate code against any VBA host (Visio, Excel, Word, AutoCAD, etc.).

## 🛠️ Generating a Custom Object Model

We provide a 2-step process to generate this model by inspecting the host application's Type Libraries (TLI).

### Step 1: Export References (VBA)
1.  **Open the VBA Editor** (`Alt+F11`) in your target application (e.g., Visio).
2.  **Import Tool**: Import `tools/VBA_Model_Exporter.bas`.
3.  **Run Export**: Execute the `ExportReferences` macro.
    *   This generates `vba_references.json` containing the GUIDs and paths of all project references.

### Step 2: Generate Model (Python)
1.  Ensure you have the Python dependencies (`pip install comtypes`).
2.  Run the generator script:
    ```bash
    python tools/generate_model.py
    ```
    *   This script uses COM to inspect the libraries and produces `vba_model.json`.
    *   It handles **CoClass member inheritance**, property name normalization, and promotes library functions (like `Mid`, `InStr`) to the global scope.
    *   **Enhancement**: It robustly extracts **Enums** and **Module-level Constants** (e.g., `visNone`, `vbCrLf`).
    *   **Type Refinement**: Automatically maps generic collection accessors (like `Selection.Item`) to specific types (e.g., `Shape`), ensuring strict validation for collection items.

### Using the Model
Pass the generated `vba_model.json` to VBAlidator using the `--model` flag:
```bash
python -m src.main ./my_code --model vba_model.json
```
VBAlidator will merge this model with the standard library to resolve symbols.

---

## 🧠 Built-in Heuristics

VBAlidator includes internal logic to handle common VBA patterns that aren't always explicitly defined in type libraries.

### 1. Form Control Dynamic Resolution
In `.frm` modules, any undefined identifier is automatically treated as an `Object` (Control). This allows code like `Me.lblStatus.Caption = "Ready"` to validate even if `lblStatus` isn't in the explicit model.

### 2. Standard VBA Callbacks
*   **`UserForm` Fallback**: If a member is not found on a Form object, VBAlidator checks the base `UserForm` class for common properties (`Show`, `Hide`, `Controls`, `Width`, `Height`).
*   **`ThisDocument` Fallback**: In projects with a `ThisDocument` module, members not found in the module are resolved against the base `Document` / `IVDocument` class.
*   **Default Member Resolution**: When an object is called like a function (e.g., `Selection(1)`), VBAlidator automatically resolves this to the object's `Item` property (e.g., `Selection.Item(1)`). This ensures that implicit collection access is strictly typed (e.g., `Selection(1)` resolves to `Shape`).

### 3. Case Insensitivity
VBA is case-insensitive. VBAlidator normalizes all model lookups (classes, members, enums, globals) to lowercase to ensure `ActivePage` matches `activepage`.

---

## 🧱 Hardcoded Metadata

### Hardcoded Keywords
The following keywords are reserved by the analyzer and cannot be used as variable names (matching VBA compiler behavior). They are hardcoded in `src/analyzer.py`:

| Category | Keywords |
| :--- | :--- |
| **Logic/Flow** | `if`, `then`, `else`, `elseif`, `end`, `exit`, `for`, `next`, `do`, `loop`, `while`, `wend`, `select`, `case`, `with`, `goto`, `resume`, `stop`, `on`, `error` |
| **Declarations** | `dim`, `static`, `const`, `public`, `private`, `global`, `sub`, `function`, `property`, `get`, `let`, `set`, `type`, `new`, `withevents`, `as` |
| **Types** | `boolean`, `integer`, `long`, `single`, `double`, `currency`, `date`, `string`, `variant`, `object`, `byte`, `decimal`, `nothing`, `empty`, `null` |
| **Operators** | `and`, `or`, `xor`, `not`, `is`, `like`, `typeof`, `mod`, `true`, `false` |
| **I/O** | `open`, `close`, `input`, `output`, `append`, `binary`, `random`, `put`, `print` |
| **Utilities** | `len`, `mid`, `redim`, `preserve`, `erase`, `call` |

### Implicit Top-Level Globals
*   **`Visio`**: Always available as an `IVApplication` proxy (if the Visio library is loaded).
*   **Forms/Classes**: Any module with `PredeclaredId = True` is automatically available as a global instance named after the class.
