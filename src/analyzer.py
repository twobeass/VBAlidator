from .parser import VariableNode, ProcedureNode, WithNode, StatementNode

class SymbolTable:
    def __init__(self, name, parent=None):
        self.name = name
        self.parent = parent
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
        self.global_scope = SymbolTable("Global")
        self.errors = []
        
        # Load Standard/Config Globals into Global Scope
        for name, defn in self.config.object_model.get("globals", {}).items():
            self.global_scope.define(name, defn.get("type", "Variant"), "Global")
            
        # Load Classes into Global Scope (as Types)
        for name in self.config.object_model.get("classes", {}).keys():
            self.global_scope.define(name, name, "Class")

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
            if mod.module_type == 'Module':
                for var in mod.variables:
                    if var.scope.lower() in ('public', 'global'):
                        self.global_scope.define(var.name, var.type_name, 'Variable')
                
                for proc in mod.procedures:
                    if proc.scope.lower() == 'public':
                         self.global_scope.define(proc.name, proc.return_type, 'Procedure')
            else:
                self.global_scope.define(mod.name, mod.name, 'Class')

    def pass2_resolution(self):
        for mod in self.modules:
            mod_scope = SymbolTable(mod.name, parent=self.global_scope)
            
            for var in mod.variables:
                mod_scope.define(var.name, var.type_name, 'Variable')
            for proc in mod.procedures:
                mod_scope.define(proc.name, proc.return_type, 'Procedure')
                
            if mod.module_type in ('Form', 'Class'):
                 mod_scope.define('Me', mod.name, 'Variable')
            
            for proc in mod.procedures:
                self.analyze_procedure(proc, mod_scope, mod)

    def analyze_procedure(self, proc, mod_scope, mod):
        proc_scope = SymbolTable(proc.name, parent=mod_scope)
        
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
            'nothing', 'empty', 'null'
        }

        iterator = iter(tokens)
        token = next(iterator, None)
        
        last_resolved_type = None
        
        while token:
            if token.type == 'IDENTIFIER':
                name = token.value
                
                if name.lower() in KEYWORDS:
                    token = next(iterator, None)
                    last_resolved_type = None # Reset chain on keyword?
                    continue

                # If we are in a dot chain: previous.name
                if last_resolved_type:
                    # Check member existence in last_resolved_type
                    member_type = self.resolve_member(last_resolved_type, name)
                    if not member_type:
                         if last_resolved_type not in ('Object', 'Variant', 'Unknown'):
                              cls_def = self.config.get_class(last_resolved_type)
                              if cls_def:
                                  if report_errors:
                                      self.errors.append({
                                          "file": filename,
                                          "line": token.line,
                                          "message": f"Member '{name}' not found in type '{last_resolved_type}' inside '{context}'."
                                      })
                              else:
                                  pass
                    last_resolved_type = member_type or 'Unknown'

                else:
                    # Root identifier
                    sym = scope.resolve(name)
                    if not sym:
                        if report_errors:
                            self.errors.append({
                                "file": filename,
                                "line": token.line,
                                "message": f"Undefined identifier '{name}' in '{context}'."
                            })
                        last_resolved_type = 'Unknown'
                    else:
                        last_resolved_type = sym['type']
            
            elif token.type == 'OPERATOR':
                if token.value == '.':
                    # Check if this is a leading dot (no last_resolved_type)
                    if last_resolved_type is None:
                        # Resolve against With Stack
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
                    # Continue chain
                    pass
                else:
                    # Other operators reset chain
                    # e.g. x + y
                    # We return the type of the last expression?
                    # For a simple static analyzer, we reset.
                    # Ideally we resolve the type of the operation result.
                    last_resolved_type = None
            
            elif token.type in ('STRING', 'INTEGER', 'FLOAT'):
                last_resolved_type = None # literal
            
            token = next(iterator, None)
        
        return last_resolved_type

    def resolve_member(self, type_name, member_name):
        cls_def = self.config.get_class(type_name)
        if not cls_def:
            return None 
        
        members = cls_def.get("members", {})
        for m_name, m_def in members.items():
            if m_name.lower() == member_name.lower():
                return m_def.get("type", "Variant")
        
        return None
