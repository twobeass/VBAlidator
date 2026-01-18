# VBA Compile Simulator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A robust, platform-independent static analysis tool for VBA (Visual Basic for Applications) code. This tool allows you to simulate the "compilation" of VBA code exported from MS Office applications on Linux/Mac environments, ensuring code integrity without opening Excel.

## 🚀 Features

*   **Static Analysis:** Detects undefined variables, invalid member access, and type mismatches.
*   **Deep Parsing:** Comprehensive parser handling nested `With` blocks, `If` statements, and loops.
*   **Conditional Compilation:** Full support for `#If...#Else` directives to simulate different environments (e.g., Win64 vs Win32).
*   **Form Support:** Parses `.frm` files to understand GUI control definitions (`TextBox`, `CommandButton`).
*   **Extensible Object Model:** Ships with standard Excel definitions and supports custom models exported from your specific Office projects.
*   **CI/CD Ready:** Returns exit codes and generates JSON reports for integration into build pipelines.

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/vba-compile-simulator.git
    cd vba-compile-simulator
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## 🛠️ Quick Start

1.  **Export your VBA code:**
    Export your modules (`.bas`), classes (`.cls`), and forms (`.frm`) to a folder (e.g., `./src-vba`).

2.  **Run the simulator:**
    ```bash
    python3 -m src.main ./src-vba
    ```

3.  **View the report:**
    Check the console output for errors or the generated `vba_report.json`.

## 📚 Documentation

Detailed documentation is available in the `docs/` folder:

*   [**Usage Guide**](docs/Usage.md): CLI arguments, options, and advanced usage.
*   [**Architecture**](docs/Architecture.md): How the Lexer, Parser, and Analyzer work.
*   [**Extending the Model**](docs/Extending.md): How to generate custom Object Models from Excel using the included VBA Exporter.

## 🧪 Testing

The repository includes a test suite to verify the analyzer against various VBA constructs.

```bash
# Run the analyzer against the test files
python3 -m src.main tests/vba_code
```

## 📄 License

This project is licensed under the MIT License.
