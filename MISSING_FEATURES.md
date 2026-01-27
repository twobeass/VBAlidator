# Missing VBA Language Features

The following VBA language features are currently not supported by the VBAlidator parser and analyzer.

## Declarations
*   **`Implements`**: Used in Class modules to implement an interface. Currently ignored or causes parsing errors.
*   **`Event`**: Used to declare user-defined events (e.g., `Public Event MyEvent(arg As String)`).
*   **`Friend`**: Scope modifier for procedures and variables (valid in Class modules). Currently ignored, causing members to be invisible.
*   **`DefType` Statements**: `DefBool`, `DefByte`, `DefInt`, `DefLong`, `DefCur`, `DefSng`, `DefDbl`, `DefDec`, `DefDate`, `DefStr`, `DefObj`, `DefVar`. These module-level default data type declarations are not parsed.

## Statements
*   **`RaiseEvent`**: Used to fire declared events. Currently treated as an undefined identifier.
*   **`GoSub` ... `Return`**: Legacy control flow statements. Treated as undefined identifiers.
*   **`LSet`**: Left-align string assignment or UDT copy. Treated as undefined identifier.
*   **`RSet`**: Right-align string assignment. Treated as undefined identifier.

## Operators
*   **`AddressOf`**: Operator used to pass procedure pointers. Currently treated as an undefined identifier.

## Directives & Options
*   **`Option Base`**: Array lower bound setting.
*   **`Option Compare`**: String comparison setting (`Binary`, `Text`, `Database`).
*   **`Option Private Module`**: Module scoping.

## Proof of Failure
Reproduction test cases have been added to `tests/vba_code/`:
*   `repro_missing_module_level.bas`
*   `repro_missing_proc_level.bas`

Running the validator on these files confirms the missing functionality via "Undefined identifier" errors or missing symbol definitions.
