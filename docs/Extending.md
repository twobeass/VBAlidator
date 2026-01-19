# Extending the Object Model

VBA heavily relies on the "Host Object Model" (e.g., Excel, Word, Outlook). Since the simulator runs outside of Office, it needs to know what objects and members exist to perform validation.

## The Object Model JSON Format

The simulator uses a JSON format to define available types.

```json
{
  "globals": {
    "ActiveSheet": { "type": "Worksheet" },
    "MsgBox": { "type": "Function", "returns": "Integer" }
  },
  "classes": {
    "Worksheet": {
      "members": {
        "Range": { "type": "Range" },
        "Name": { "type": "String" }
      }
    },
    "Range": {
      "members": {
        "Value": { "type": "Variant" }
      }
    }
  }
}
```

- **globals**: Objects or functions available everywhere (e.g., `Application`, `ActiveSheet`).
- **classes**: Type definitions detailing available members (properties/methods).

## Generating a Model from Office

Instead of writing this JSON manually, the project includes a **VBA Exporter** tool.

### Using `VBA_Model_Exporter.bas`

1.  Open your Excel/Word project.
2.  Import `VBA_Model_Exporter.bas` into the VBE (Visual Basic Editor).
3.  Ensure you have **"Trust access to the VBA project object model"** enabled in Excel Trust Center settings.
4.  Run the `ExportModel` macro.

The macro will:
- Attempt to use the `TypeLib Information` (TLI) library if available to export detailed definitions of all referenced libraries.
- Fallback to a basic export if TLI is not available.
- Generate a `vba_model.json` file in the same directory as your workbook.

### Loading the Custom Model

Once you have `vba_model.json`, pass it to the simulator:

```bash
python3 -m src.main ./my_code --model ./vba_model.json
```

## Code-Level Heuristics

In addition to the JSON model, the analyzer implements several **heuristics** in `src/analyzer.py` to handle dynamic or common VBA patterns without requiring explicit definitions:

1.  **Prefix-Based Resolution**: Identifiers starting with specific prefixes are automatically resolved to a type.
    *   `txt*`, `lbl*`, `btn*`... -> `Control` (Useful for Forms)
    *   `vb*`, `mso*`, `ad*` -> `Long` (Intrinsic Constants)
    *   `vis*` -> `Long` (Visio Constants, excluding `Visio` the object)

2.  **Common Member Resolution**: A list of common identifiers (e.g., `Name`, `Count`, `Item`, `Add`, `Close`) are resolved to `Variant` if not found in the explicit model. This acts as a catch-all to prevent "Member not found" errors for dynamic objects.

3.  **Implicit Globals**: The analyzer automatically registers `Visio` as a global `Application` object and creates implicit global instances for Forms and Classes with `VB_PredeclaredId = True`.
