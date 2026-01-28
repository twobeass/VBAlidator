# VBAlidator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**VBAlidator** is a robust, platform-independent static analysis tool for VBA (Visual Basic for Applications). It allows you to simulate the "compilation" of VBA code exported from MS Office applications on any environment (Windows, Linux, or Mac), ensuring code integrity without needing to open the Office application itself.

## 🚀 Features

*   **Static Analysis:** Detects undefined variables, invalid member access, and type mismatches.
*   **Dynamic Object Model:** Generate and load object models for *any* VBA host (Visio, Excel, Word, AutoCAD) using the integrated TLI-based generator.
*   **Deep Parsing:** Comprehensive parser handling nested `With` blocks, multi-statement lines, and complex loops.
*   **Conditional Compilation:** Full support for `#If...#Else` directives to simulate different environments (e.g., Win64 vs Win32).
*   **Form Support:** Intelligent heuristics for `.frm` files to handle implicit GUI controls and `UserForm` members.
*   **CI/CD Ready:** Returns exit codes and generates JSON reports for seamless integration into build pipelines.

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/twobeass/VBAlidator.git
    cd VBAlidator
    ```

2.  **Install the package:**
    ```bash
    pip install .
    ```
    This will install the dependencies and the `vbalidator` command-line tool.

## 🛠️ Quick Start

### 1. Generate the Object Model
To validate host-specific code (e.g., `Visio.Shape`), you first need a model of that host.
1. Use `tools/VBA_Model_Exporter.bas` in your Office application to export references.
2. Run the generator: `python tools/generate_model.py`.

### 2. Run the Validator
If you have a `vba_model.json` in your current directory, it will be automatically used.
```bash
vbalidator ./path/to/vba_code
```

You can also specify the model manually:
```bash
vbalidator ./path/to/vba_code --model /path/to/vba_model.json
```

## 📚 Documentation

Detailed documentation is available in the `docs/` folder:

*   [**Usage Guide**](docs/Usage.md): CLI arguments, options, and advanced usage.
*   [**Configuration**](docs/Configuration.md): How to generate custom Object Models and how heuristics work.
*   [**Architecture**](docs/Architecture.md): Deep dive into the Lexer, Parser, and Analyzer.

## 🧪 Testing

```bash
vbalidator tests/samples/valid_code
```

To see the tool in action against intentional errors, check the `tests/demo` folder:
```bash
vbalidator tests/demo
```

## 📄 License

This project is licensed under the MIT License.
