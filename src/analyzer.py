from .parser import VariableNode, ProcedureNode, WithNode, StatementNode

class SymbolTable:
    def __init__(self, name, parent=None, scope_type='Block'):
        self.name = name
        self.parent = parent
        self.scope_type = scope_type
        self.symbols = {} # name -> {type: ..., kind: Var/Proc/Class}

    def define(self, name, type_name, kind):
        self.symbols[name.lower()] = {"type": type_name, "kind": kind}

    def resolve(self, name):
        name = name.lower()
        if name in self.symbols:
            return self.symbols[name]
        if self.parent:
            return self.parent.resolve(name)
        return None

class Analyzer:
    def __init__(self, config):
        self.config = config
        self.modules = []
        self.config = config
        self.modules = []
        self.global_scope = SymbolTable("Global", scope_type='Global')
        self.errors = []
        self.udts = {} # name_lower -> TypeNode
        
        # Load Standard/Config Globals into Global Scope
        for name, defn in self.config.object_model.get("globals", {}).items():
            self.global_scope.define(name, defn.get("type", "Variant"), "Global")
            
        # Load Classes into Global Scope (as Types)
        for name in self.config.object_model.get("classes", {}).keys():
            self.global_scope.define(name, name, "Class")

        # Load References as Global Symbols (Treat as Objects/Libraries)
        # Allows usage like Visio.Application or Excel.Range
        if "references" in self.config.object_model:
            for ref in self.config.object_model["references"]:
                 # print(f"DEBUG: Registering Library {ref['name']}")
                 self.global_scope.define(ref["name"], "Object", "Library")
        
        # Manual Visio Alias
        self.global_scope.define("Visio", "Application", "Library")
        
        # Verify Visio registration
        # res = self.global_scope.resolve("Visio")
        # if not res: print("DEBUG WARNING: Visio not found in global scope after init")
        # else: print(f"DEBUG: Visio resolved to {res}")

    def add_module(self, module_node):
        self.modules.append(module_node)

    def analyze(self):
        # Pass 1: Populate Symbol Tables
        self.pass1_discovery()
        
        # Pass 2: Verify References
        self.pass2_resolution()
        
        return self.errors

    def pass1_discovery(self):
        for mod in self.modules:
            # Register module name itself (allows usage like Module1.Func)
            # Use module name as type so resolve_member finds it
            self.global_scope.define(mod.name, mod.name, mod.module_type)
            
            # Check for Predeclared ID (Global Instance for Classes/Forms)
            if mod.attributes.get('VB_PredeclaredId', 'False').lower() == 'true':
                 # Implicit global variable with same name as class
                 self.global_scope.define(mod.name, mod.name, mod.module_type)

            if mod.module_type == 'Module':
                for var in mod.variables:
                    if var.scope.lower() in ('public', 'global'):
                        self.global_scope.define(var.name, var.type_name, 'Variable')
                
                for proc in mod.procedures:
                    if proc.scope.lower() == 'public':
                         self.global_scope.define(proc.name, proc.return_type, 'Procedure')
                
                # Register Public Types
                for type_name, udt in mod.types.items():
                    if udt.scope.lower() == 'public':
                        self.global_scope.define(type_name, type_name, 'Type')
                        self.udts[type_name.lower()] = udt

            else:
                self.global_scope.define(mod.name, mod.name, 'Class')
                # Also register types in classes/forms? Usually Private but can be Public?
                for type_name, udt in mod.types.items():
                     # Even if private, we might need to store them for module-level resolution?
                     # Global scope only needs Public.
                     if udt.scope.lower() == 'public':
                         self.global_scope.define(type_name, type_name, 'Type')
                         self.udts[type_name.lower()] = udt

    def pass2_resolution(self):
        for mod in self.modules:
            mod_scope = SymbolTable(mod.name, parent=self.global_scope, scope_type=mod.module_type)
            
            for var in mod.variables:
                mod_scope.define(var.name, var.type_name, 'Variable')
            for proc in mod.procedures:
                mod_scope.define(proc.name, proc.return_type, 'Procedure')
            # Register Local/Private Types in Module Scope
            for type_name, udt in mod.types.items():
                mod_scope.define(type_name, type_name, 'Type')
                self.udts[type_name.lower()] = udt
                
            if mod.module_type in ('Form', 'Class'):
                 mod_scope.define('Me', mod.name, 'Variable')
            
            for proc in mod.procedures:
                self.analyze_procedure(proc, mod_scope, mod)

    def analyze_procedure(self, proc, mod_scope, mod):
        proc_scope = SymbolTable(proc.name, parent=mod_scope, scope_type='Procedure')
        
        for arg in proc.args:
            proc_scope.define(arg.name, arg.type_name, 'Variable')
            
        # Locals are now in the body as statements, we must parse them on the fly
        # recursive block analysis
        self.analyze_block(proc.body, proc_scope, mod.filename, proc.name, with_stack=[])

    def analyze_block(self, nodes, scope, filename, context, with_stack):
        for node in nodes:
            if isinstance(node, StatementNode):
                # Check for Dim
                if node.tokens and node.tokens[0].value.lower() in ('dim', 'static'):
                     self.process_dim(node.tokens, scope)
                else:
                     self.analyze_statement(node.tokens, scope, filename, context, with_stack)
            
            elif isinstance(node, WithNode):
                # Resolve expression type
                expr_type = self.resolve_expression_type(node.expr_tokens, scope, with_stack)
                # Push to stack
                # If unknown, we push 'Unknown' to suppress errors inside?
                # Or we warn?
                # We push whatever we found.
                new_stack = with_stack + [expr_type or 'Unknown']
                self.analyze_block(node.body, scope, filename, context, new_stack)

    def process_dim(self, tokens, scope):
        # Extremely simplified Dim parser to populate scope
        # Dim x As Integer, y As String
        # We need to skip 'Dim'
        iterator = iter(tokens)
        next(iterator) # Skip Dim
        
        # We assume valid syntax from Parser, just extract names and types
        current_name = None
        current_type = 'Variant'
        
        tokens_list = list(iterator)
        i = 0
        while i < len(tokens_list):
            t = tokens_list[i]
            if t.type == 'IDENTIFIER':
                if t.value.lower() == 'as':
                    # Parse type
                    i += 1
                    type_parts = []
                    while i < len(tokens_list):
                        # Handle New
                        if tokens_list[i].value.lower() == 'new':
                            i += 1
                            continue
                        
                        if tokens_list[i].type == 'IDENTIFIER':
                            type_parts.append(tokens_list[i].value)
                            i += 1
                            if i < len(tokens_list) and tokens_list[i].value == '.':
                                type_parts.append('.')
                                i += 1
                            else:
                                break
                        else:
                            break
                    current_type = "".join(type_parts)
                    # Register previous var
                    if current_name:
                        scope.define(current_name, current_type, 'Variable')
                        current_name = None
                        current_type = 'Variant'
                    else:
                        pass # 'As' without Name? Error
                else:
                    if current_name:
                        # previous was variant
                        scope.define(current_name, 'Variant', 'Variable')
                    current_name = t.value
                    i += 1
            elif t.value == ',':
                if current_name:
                    scope.define(current_name, current_type, 'Variable')
                    current_name = None
                    current_type = 'Variant'
                i += 1
            else:
                i += 1
        
        if current_name:
             scope.define(current_name, current_type, 'Variable')

    def resolve_expression_type(self, tokens, scope, with_stack):
        # Reuse analyze_statement logic but return the final type
        # For 'ActiveSheet.Range("A1")', we want 'Range'.
        # We scan the tokens and track type.
        
        # Simplified: Just run logic and return last resolved type
        return self.analyze_statement(tokens, scope, "", "", with_stack, report_errors=False)

    def analyze_statement(self, tokens, scope, filename, context, with_stack, report_errors=True):
        KEYWORDS = {
            'set', 'call', 'if', 'then', 'else', 'elseif', 'end', 'exit', 
            'on', 'error', 'goto', 'resume', 'do', 'loop', 'while', 'wend', 
            'for', 'next', 'select', 'case', 'with', 'to', 'step', 'in',
            'byval', 'byref', 'optional', 'paramarray', 'true', 'false',
            'nothing', 'empty', 'null',
            'not', 'each', 'sub', 'function', 'property', 'const', 'dim', 'as', 
            'type', 'boolean', 'integer', 'string', 'variant', 'object', 
            'byte', 'long', 'single', 'double', 'currency', 'date', 'decimal',
            'and', 'or', 'xor', 'is', 'like', 'typeof', 'mod', 'new', 'print',
            'open', 'close', 'input', 'output', 'append', 'binary', 'random',
            'get', 'put', 'let', 'stop', 'len', 'mid', 'redim', 'preserve', 'erase'
        }

        # Use index based loop for peek ability
        i = 0
        last_resolved_type = None
        expect_member = False
        prev_keyword = None
        
        while i < len(tokens):
            token = tokens[i]
            
            if token.type == 'IDENTIFIER':
                name = token.value
                
                if name.lower() in KEYWORDS and not expect_member:
                    prev_keyword = name.lower()
                    last_resolved_type = None
                    i += 1
                    continue
                
                # Check for Label Definition "Label:"
                if i + 1 < len(tokens) and tokens[i+1].type == 'OPERATOR' and tokens[i+1].value == ':':
                    # Label definition - skip identifier and colon
                    i += 2
                    prev_keyword = None
                    last_resolved_type = None
                    continue

                # Check for GoTo/Resume Label skipping
                if prev_keyword in ('goto', 'resume'):
                    prev_keyword = None
                    i += 1
                    continue

                # Dot chain
                if expect_member and last_resolved_type:
                    member_type = self.resolve_member(last_resolved_type, name)
                    if not member_type:
                        if last_resolved_type not in ('Object', 'Variant', 'Unknown', 'Control', 'Form'):
                            if report_errors:
                                self.errors.append({
                                    "file": filename,
                                    "line": token.line,
                                    "message": f"Member '{name}' not found in type '{last_resolved_type}' inside '{context}'."
                                })
                    last_resolved_type = member_type or 'Unknown'
                    expect_member = False
                else:
                    # Root identifier
                    sym = scope.resolve(name)
                    if not sym:
                        # Heuristic: Auto-resolve constants (vb*, vis*, mso*)
                        low_name = name.lower()
                        if low_name == 'visio':
                             last_resolved_type = 'Application'
                        elif (low_name.startswith('vb') or low_name.startswith('vis') or low_name.startswith('mso') or low_name.startswith('ad')) and low_name != 'visio':
                             last_resolved_type = 'Long'
                        
                        # Heuristic: Form Controls (if in Form/Class context and unknown)
                        # Heuristic: Form Controls (if in Form/Class context and unknown)
                        elif scope.scope_type in ('Form', 'Class') or (scope.parent and scope.parent.scope_type in ('Form', 'Class')):
                             last_resolved_type = 'Control'

                        # Heuristic: Common VBA functions if missing from model
                        # Heuristic: Common VBA functions if missing from model
                        # Heuristic: Common VBA functions if missing from model
                        # Heuristic: Common VBA functions if missing from model
                        elif low_name in ('instr', 'left', 'right', 'mid', 'len', 'replace', 'chr', 'format', 'timer', 'doevents', 'clng', 'cint', 'cdbl', 'cstr', 'ubound', 'lbound', 'dir', 'curdir', 'kill', 'split', 'join', 'array', 'isnumeric', 'isempty', 'isnothing', 'isobject', 'isdate', 'date', 'now', 'space', 'string', 'activewindow', 'activedocument', 'activepage', 'userforms', 'load', 'unload', 'redim', 'erase'):
                             last_resolved_type = 'Variant'
                        # Heuristic: Form Controls (Prefixes)
                        elif low_name.startswith(('txt', 'lbl', 'cmd', 'btn', 'lst', 'opt', 'chk', 'img', 'fra')):
                             last_resolved_type = 'Control'

                        # Heuristic: VBA library prefix
                        elif low_name == 'vba':
                             last_resolved_type = 'VBA' # Pseudo-lib
                        else:
                            if report_errors:
                                # DEBUG
                                if name.lower() == 'lstmodelle':
                                     print(f"DEBUG: lstModelle scope_type={getattr(scope, 'scope_type', 'N/A')} parent_type={getattr(scope.parent, 'scope_type', 'N/A') if scope.parent else 'None'}")
                                
                                self.errors.append({
                                    "file": filename,
                                    "line": token.line,
                                    "message": f"Undefined identifier '{name}' in '{context}'."
                                })
                            last_resolved_type = 'Unknown'
                    else:
                        last_resolved_type = sym['type']
                
                prev_keyword = None
                i += 1
            
            elif token.type == 'OPERATOR':
                if token.value == '.':
                    if last_resolved_type is None:
                        if with_stack:
                            last_resolved_type = with_stack[-1]
                        else:
                            if report_errors:
                                self.errors.append({
                                    "file": filename,
                                    "line": token.line,
                                    "message": f"Invalid or unexpected . reference without With block in '{context}'."
                                })
                            last_resolved_type = 'Unknown'
                    expect_member = True
                    i += 1
                elif token.value == '(':
                    # Function call or array access: A(1).B
                    # Skip until matching )
                    if last_resolved_type:
                        depth = 1
                        i += 1
                        while i < len(tokens) and depth > 0:
                            if tokens[i].type == 'OPERATOR':
                                if tokens[i].value == '(': depth += 1
                                elif tokens[i].value == ')': depth -= 1
                            i += 1
                        # After call/index, result is likely Variant/Object
                        last_resolved_type = 'Variant'
                        expect_member = False
                    else:
                        last_resolved_type = None
                        expect_member = False
                        prev_keyword = None
                        i += 1
                else:
                    last_resolved_type = None
                    expect_member = False
                    prev_keyword = None
                    i += 1
            
            elif token.type in ('STRING', 'INTEGER', 'FLOAT'):
                last_resolved_type = None
                expect_member = False
                prev_keyword = None
                i += 1
            else:
                i += 1
        
        return last_resolved_type

    def resolve_member(self, type_name, member_name):
        return self._resolve_member_internal(type_name, member_name, try_aliases=True)

    def _resolve_member_internal(self, type_name, member_name, try_aliases=True):
        # Handle Qualified Types (e.g. Visio.Document -> Document)
        if '.' in type_name:
             simple_name = type_name.split('.')[-1]
             # Try resolving full name first
             res = self._resolve_member_base(type_name, member_name)
             if res: return res
             # Then simple name
             res = self._resolve_member_base(simple_name, member_name)
             if res: return res
             # If aliases enabled, try simple name aliases
             if try_aliases:
                 return self._resolve_with_aliases(simple_name, member_name)
             return None
        
        # Non-qualified
        res = self._resolve_member_base(type_name, member_name)
        if res: return res
        
        if try_aliases:
             return self._resolve_with_aliases(type_name, member_name)
        return None

    def _resolve_with_aliases(self, type_name, member_name):
        # Visio Interface Mapping (CoClass -> Interface)
        ALIASES = {
            "Document": "IVDocument",
            "Page": "IVPage",
            "Master": "IVMaster",
            "Shape": "IVShape",
            "Cell": "IVCell",
            "Selection": "IVSelection",
            "Window": "IVWindow",
            "Application": "IVApplication",
            "Hyperlink": "IVHyperlink",
            "Connect": "IVConnect",
            "Layer": "IVLayer",
            "Style": "IVStyle",
            "Layer": "IVLayer",
            "Style": "IVStyle",
            "Font": "IVFont",
            "Event": "IVEvent",
            "ThisDocument": "IVDocument",
            "Documents": "IVDocuments",
            "Pages": "IVPages",
            "Shapes": "IVShapes",
            "Windows": "IVWindows",
            "Masters": "IVMasters",
            "Connects": "IVConnects"
        }
        if type_name in ALIASES:
            return self._resolve_member_base(ALIASES[type_name], member_name)
        return None

        return None

    def _resolve_member_base(self, type_name, member_name):
        # 0. Check UDTs
        if type_name.lower() in self.udts:
            udt = self.udts[type_name.lower()]
            for m in udt.members:
                if m.name.lower() == member_name.lower():
                    return m.type_name
        
        # 1. Check Config Classes
        cls_def = self.config.get_class(type_name)
        if cls_def:
            members = cls_def.get("members", {})
            for m_name, m_def in members.items():
                if m_name.lower() == member_name.lower():
                    return m_def.get("type", "Variant")
        
        # 2. Check Standard Modules (if type_name matches a module name)
        for mod in self.modules:
            if mod.module_type == 'Module' and mod.name.lower() == type_name.lower():
                # Check public vars/procs
                for v in mod.variables:
                    if v.name.lower() == member_name.lower() and v.scope.lower() in ('public', 'global'):
                        return v.type_name
                for p in mod.procedures:
                    if p.name.lower() == member_name.lower() and p.scope.lower() == 'public':
                         return p.return_type
        
        # 3. Check Project Classes (Class Modules)
        for mod in self.modules:
            if mod.module_type in ('Class', 'Form') and mod.name.lower() == type_name.lower():
                 for v in mod.variables:
                     if v.name.lower() == member_name.lower() and v.scope.lower() in ('public', 'global'):
                         return v.type_name
                 for p in mod.procedures:
                     if p.name.lower() == member_name.lower() and p.scope.lower() == 'public':
                         return p.return_type

        # 4. Check Pseudo-LIBS and Heuristic Constants
        if type_name == 'VBA':
             return "Variant"
        
        # Heuristic: Allow vb* and vis* as members of anything (or Application/Visio)
        if member_name.lower().startswith('vis') or member_name.lower().startswith('vb'):
            return "Long"

        # Heuristic: Common Visio/Excel members
        # Heuristic: Form Controls (Prefixes)
        if member_name.lower().startswith(('txt', 'lbl', 'cmd', 'btn', 'lst', 'opt', 'chk', 'img', 'fra')):
             return "Control"

        # Heuristic: Common Visio/Excel members
        if member_name.lower() in ('cellsu', 'cellexistsu', 'document', 'application', 'stat', 'objecttype', 'master', 'shapes', 'pages', 'nameu', 'text', 'count', 'item', 'one', 'activedocument', 'activewindow', 'activepage', 'selection', 'containingpage', 'diagramservicesenabled', 'shape', 'parent', 'id', 'show', 'hide', 'caption', 'tag', 'refresh', 'datacolumns', 'getprimarykey', 'getdatarowids', 'getrowdata', 'selecteddatarecordset', 'selecteddatarowid', 'list', 'listindex', 'additem', 'clear', 'value', 'exists', 'keys', 'remove', 'add', 'description', 'name', 'enabled', 'controls', 'cells', 'addnamedrow', 'listcount', 'selected', 'width', 'height', 'characters', 'charcount', 'copy', 'cellssrc', 'oned', 'section', 'connects', 'sectionexists', 'resultstr', 'rowtype', 'xytopage', 'cellexists', 'left', 'top', 'columncount', 'scrollheight', 'columnwidths', 'uniqueid', 'spatialrelation', 'boundingbox', 'containingshape', 'pagesheet', 'memberofcontainers', 'drop', 'nameid', 'hyperlinks', 'delete', 'replaceshape', 'addsection', 'addrow', 'addhyperlink', 'row', 'formulau', 'breaklinktodata', 'linktodata', 'deleterow', 'rowcount', 'intersect', 'type'):
             return "Variant"

        return None
