# Architecture Overview

The **VBA Compile Simulator** is built as a modular static analysis pipeline using Python. It avoids regex-only parsing in favor of a proper tokenization and parsing strategy to handle nested scopes and complex VBA syntax.

## Core Components

### 1. Lexer (`src/lexer.py`)
*   **Responsibility:** Converts raw VBA source code into a stream of `Token` objects.
*   **Features:**
    *   Handles comments, strings, identifiers, and literals.
    *   Recognizes line continuations (` _`).
    *   Identifies Preprocessor directives (`#If`, `#Const`, `#End`).

### 2. Preprocessor (`src/preprocessor.py`)
*   **Responsibility:** Filters the token stream based on conditional compilation logic.
*   **Logic:**
    *   Maintains a stack of active/inactive states based on `#If...#Else` blocks.
    *   Evaluates boolean expressions using the definitions provided via `--define` or `#Const` directives within the code.
    *   Strips out code that would be inactive in the target environment.

### 3. Parser (`src/parser.py`)
*   **Responsibility:** Consumes tokens to build an Abstract Syntax Tree (AST) or structured Node representation.
*   **Components:**
    *   **FormParser:** Extracts GUI control definitions from `.frm` headers.
    *   **VBAParser:** Parses the actual code logic.
*   **Key Structures:**
    *   `ModuleNode`: Represents a file (Module, Class, Form).
    *   `ProcedureNode`: Represents a Sub, Function, or Property.
    *   `StatementNode`: Represents a single line of code.
    *   `WithNode`: Represents a `With...End With` block (recursive).

### 4. Analyzer (`src/analyzer.py`)
*   **Responsibility:** Performs semantic analysis on the parsed nodes.
*   **Process:**
    *   **Pass 1 (Discovery):** Scans all modules to register global variables, public procedures, and class names into the Global Symbol Table.
    *   **Pass 2 (Resolution):** Walks through procedure bodies to verify logic.
        *   Creates local scopes for arguments and `Dim` variables.
        *   Maintains a `With Stack` to resolve dot-notation (`.Value`) against the current `With` object context.
        *   Validates member access against the loaded Object Model.

### 5. Configuration (`src/config.py`)
*   **Responsibility:** Manages the Object Model definitions.
*   **Data:** Loads `src/std_model.json` by default and merges any user-provided JSON models.

## Data Flow

```
Source Code (.bas) -> Lexer -> Tokens -> Preprocessor -> Filtered Tokens 
                                                              |
                                                              v
AST Nodes <----------------------------------------------- Parser
    |
    v
Analyzer (Pass 1 & 2) -> Symbol Tables -> Error Report -> CLI Output
```
