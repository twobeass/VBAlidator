import argparse
import os
import sys
import json
from colorama import init, Fore, Style

from .config import Config
from .lexer import Lexer
from .preprocessor import Preprocessor
from .parser import VBAParser, FormParser, ModuleNode
from .analyzer import Analyzer

init(autoreset=True)

def main():
    parser = argparse.ArgumentParser(description="VBAlidator - VBA Static Analysis Tool")
    parser.add_argument("input_folder", help="Path to the folder containing VBA files")
    parser.add_argument("--define", help="Conditional compilation constants (e.g. 'WIN64=True,VBA7=True')")
    parser.add_argument("--model", help="Path to a custom JSON object model definition file")
    parser.add_argument("--output", help="Path to save the JSON report", default="vba_report.json")
    
    args = parser.parse_args()
    
    # 1. Configuration
    config = Config()
    if args.define:
        config.parse_defines(args.define)

    try:
        if args.model:
            config.load_model(args.model)
        elif os.path.exists("vba_model.json"):
            print(Fore.CYAN + "Loading implicit model: vba_model.json")
            config.load_model("vba_model.json")
    except Exception as e:
        print(Fore.RED + f"Error loading model: {e}")
        sys.exit(1)

    analyzer = Analyzer(config)
    
    # 2. File Discovery
    if not os.path.exists(args.input_folder):
        print(Fore.RED + f"Error: Input folder '{args.input_folder}' does not exist.")
        sys.exit(1)

    files = []
    for root, _, filenames in os.walk(args.input_folder):
        for f in filenames:
            if f.lower().endswith(('.cls', '.bas', '.frm')):
                files.append(os.path.join(root, f))
                
    print(Fore.CYAN + f"Found {len(files)} VBA files in {args.input_folder} and subdirectories")

    # 3. Processing Loop
    for filepath in files:
        filename = os.path.relpath(filepath, args.input_folder)
        try:
            with open(filepath, 'r', encoding='latin-1') as f: # VBA export is often latin-1 or cp1252
                content = f.read()
            
            # Determine module type
            ext = os.path.splitext(filename)[1].lower()
            module_type = 'Module'
            if ext == '.cls': module_type = 'Class'
            elif ext == '.frm': module_type = 'Form'
            
            # Form Handling
            controls = []
            code_content = content
            if ext == '.frm':
                fp = FormParser()
                controls = fp.parse(content)
                # Find where code attributes start to skip Form definition header
                import re
                # Find the start of Attributes
                match = re.search(r'Attribute\s+VB_Name', content)
                if match:
                    code_content = content[match.start():]
            
            # Lexer
            lexer = Lexer(code_content)
            tokens = list(lexer.tokenize())
            
            # Preprocessor
            pp = Preprocessor(tokens, config.definitions)
            processed_tokens = list(pp.process())
            
            # Parser
            parser = VBAParser(processed_tokens, filename=filename)
            module_node = parser.parse_module()
            module_node.filename = filename
            module_node.module_type = module_type
            
            # Add controls to module variables if Form
            if ext == '.frm':
                module_node.variables.extend(controls)
                
            analyzer.add_module(module_node)
            
        except Exception as e:
            print(Fore.RED + f"Error processing {filename}: {e}")
            import traceback
            traceback.print_exc()

    # 4. Analysis
    print(Fore.YELLOW + "Analyzing...")
    errors = analyzer.analyze()

    # 5. Reporting
    print(Fore.GREEN + f"Analysis Complete. Found {len(errors)} potential issues.")
    
    report_data = {
        "summary": {
            "files_scanned": len(files),
            "issues_found": len(errors)
        },
        "issues": errors
    }
    
    # Console Output
    for err in errors:
        print(f"{Fore.MAGENTA}{err['file']}:{err['line']}: {Fore.RED}{err['message']}")

    # JSON Output
    with open(args.output, 'w') as f:
        json.dump(report_data, f, indent=2)
    print(Fore.CYAN + f"Report saved to {args.output}")

if __name__ == "__main__":
    main()
