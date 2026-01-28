from .parser import VariableNode, ProcedureNode, WithNode, StatementNode

class SymbolTable:
    def __init__(self, name, parent=None, scope_type='Block'):
        self.name = name
        self.parent = parent
        self.scope_type = scope_type
        self.symbols = {} # name -> {type: ..., kind: Var/Proc/Class, extra: ...}

    def define(self, name, type_name, kind, extra=None):
        self.symbols[name.lower()] = {"type": type_name, "kind": kind, "extra": extra}

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
        self.global_scope = SymbolTable("Global", scope_type='Global')
        self.errors = []
        self.udts = {} # name_lower -> TypeNode
        
        # Load Standard/Config Globals into Global Scope
        for name, defn in self.config.object_model.get("globals", {}).items():
            # Use 'returns' as type if available, otherwise 'type'
            type_name = defn.get("returns", defn.get("type", "Variant"))
            kind = defn.get("type", "Global")
            self.global_scope.define(name, type_name, kind, extra=defn)
            
        # Load Classes into Global Scope (as Types)
        for name in self.config.object_model.get("classes", {}).keys():
            self.global_scope.define(name, name, "Class")

        # Load References as Global Symbols (Treat as Objects/Libraries)
        if "references" in self.config.object_model:
            for ref in self.config.object_model["references"]:
                 self.global_scope.define(ref["name"], "Object", "Library")

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
            self.global_scope.define(mod.name, mod.name, mod.module_type)
            
            # Check for Predeclared ID (Global Instance for Classes/Forms)
            if mod.attributes.get('VB_PredeclaredId', 'False').lower() == 'true':
                 self.global_scope.define(mod.name, mod.name, mod.module_type)

            if mod.module_type == 'Module':
                for var in mod.variables:
                    if var.scope.lower() in ('public', 'global', 'friend'):
                        self.global_scope.define(var.name, var.type_name, 'Variable')
                
                for proc in mod.procedures:
                    if proc.scope.lower() in ('public', 'friend'):
                         self.global_scope.define(proc.name, proc.return_type, 'Procedure', extra=proc)
                
                # Register Public Types
                for type_name, udt in mod.types.items():
                    if udt.scope.lower() in ('public', 'friend'):
                        self.global_scope.define(type_name, type_name, 'Type')
                        self.udts[type_name.lower()] = udt

            else:
                self.global_scope.define(mod.name, mod.name, 'Class')
                for type_name, udt in mod.types.items():
                     if udt.scope.lower() in ('public', 'friend'):
                         self.global_scope.define(type_name, type_name, 'Type')
                         self.udts[type_name.lower()] = udt

    def pass2_resolution(self):
        for mod in self.modules:
            mod_scope = SymbolTable(mod.name, parent=self.global_scope, scope_type=mod.module_type)
            
            for var in mod.variables:
                mod_scope.define(var.name, var.type_name, 'Variable')
            for proc in mod.procedures:
                mod_scope.define(proc.name, proc.return_type, 'Procedure', extra=proc)
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
            
        self.analyze_block(proc.body, proc_scope, mod.filename, proc.name, with_stack=[])

    def analyze_block(self, nodes, scope, filename, context, with_stack):
        unreachable = False
        prev_node = None

        for node in nodes:
            if isinstance(node, StatementNode):
                # Check for Label
                if self.is_label(node.tokens):
                    unreachable = False

                # Check for Control Flow Boundary (e.g. End If, Else, Next) -> Reset unreachable
                # Heuristic: If we hit a block boundary, assume the jump was conditional or we merged back
                if unreachable and self.is_control_flow_boundary(node.tokens):
                    unreachable = False

                if unreachable:
                    if not self.is_ignorable(node.tokens):
                        self.errors.append({
                            "file": filename,
                            "line": node.tokens[0].line,
                            "message": f"Unreachable code detected in '{context}'."
                        })

                # Check for Dim
                if node.tokens and node.tokens[0].value.lower() in ('dim', 'static', 'const'):
                     self.process_dim(node.tokens, scope, filename, context, with_stack)
                else:
                     self.analyze_statement(node.tokens, scope, filename, context, with_stack)

                # Check for Jump
                if self.is_unconditional_jump(node.tokens):
                    # Check if conditional (e.g. "If x Then Exit Sub" split by colon)
                    is_conditional_jump = False
                    if prev_node and isinstance(prev_node, StatementNode):
                         # Check if on same line
                         if prev_node.tokens and node.tokens and prev_node.tokens[0].line == node.tokens[0].line:
                             # Check if prev starts with If
                             if prev_node.tokens[0].value.lower() == 'if':
                                 is_conditional_jump = True

                    if not is_conditional_jump:
                        unreachable = True
            
            elif isinstance(node, WithNode):
                if unreachable:
                     self.errors.append({
                         "file": filename,
                         "line": node.expr_tokens[0].line if node.expr_tokens else 0,
                         "message": f"Unreachable code detected (With block) in '{context}'."
                     })

                expr_type = self.resolve_expression_type(node.expr_tokens, scope, with_stack)
                new_stack = with_stack + [expr_type or 'Unknown']
                self.analyze_block(node.body, scope, filename, context, new_stack)

            prev_node = node

    def process_dim(self, tokens, scope, filename, context, with_stack):
        # Simplified Dim parser
        iterator = iter(tokens)
        next(iterator) # Skip Dim
        
        current_name = None
        current_type = 'Variant'
        is_array = False
        
        tokens_list = list(iterator)
        i = 0
        while i < len(tokens_list):
            t = tokens_list[i]
            if t.type == 'IDENTIFIER':
                if t.value.lower() == 'as':
                    i += 1
                    type_parts = []
                    while i < len(tokens_list):
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
                    if is_array:
                        current_type += "()"

                    if current_name:
                        if current_name.lower() in scope.symbols:
                            self.errors.append({
                                "file": filename,
                                "line": tokens[0].line,
                                "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                            })
                        else:
                            scope.define(current_name, current_type, 'Variable')
                        current_name = None
                        current_type = 'Variant'
                        is_array = False
                else:
                    if current_name:
                        # Implicit Variant definition for the previous variable
                        t_type = "Variant"
                        if is_array: t_type += "()"
                        scope.define(current_name, t_type, 'Variable')
                        is_array = False # Reset for next

                    current_name = t.value
                    i += 1
                    
                    # Check for Array ()
                    if i < len(tokens_list) and tokens_list[i].value == '(':
                        is_array = True
                        depth = 1
                        i += 1
                        while i < len(tokens_list) and depth > 0:
                            if tokens_list[i].value == '(': depth += 1
                            elif tokens_list[i].value == ')': depth -= 1
                            i += 1

            elif t.type == 'OPERATOR' and t.value == '=':
                # Initialization
                if current_name:
                     if current_name.lower() in scope.symbols:
                         self.errors.append({
                             "file": filename,
                             "line": tokens[0].line,
                             "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                         })
                     else:
                         t_type = current_type
                         if is_array and not t_type.endswith('()'): t_type += "()"
                         scope.define(current_name, t_type, 'Variable')

                     # We defined the variable, now let's analyze the assignment expression
                     # We need to find where the expression ends (at comma or end of tokens)
                     expr_start = i + 1
                     expr_end = len(tokens_list)

                     # Look ahead for comma
                     depth_parens = 0
                     for k in range(expr_start, len(tokens_list)):
                          if tokens_list[k].value == '(': depth_parens += 1
                          elif tokens_list[k].value == ')': depth_parens -= 1
                          elif tokens_list[k].value == ',' and depth_parens == 0:
                               expr_end = k
                               break

                     expr_tokens = tokens_list[expr_start:expr_end]
                     self.analyze_statement(expr_tokens, scope, filename, context, with_stack)

                     current_name = None
                     current_type = 'Variant'
                     is_array = False

                     i = expr_end
                else:
                     i += 1

            elif t.value == ',':
                if current_name:
                    if current_name.lower() in scope.symbols:
                        self.errors.append({
                            "file": filename, 
                            "line": tokens[0].line,
                            "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                        })
                    else:
                        t_type = current_type
                        if is_array and not t_type.endswith('()'): t_type += "()"
                        scope.define(current_name, t_type, 'Variable')
                    current_name = None
                    current_type = 'Variant'
                    is_array = False
                i += 1
            else:
                i += 1
        
        if current_name:
             if current_name.lower() in scope.symbols:
                 self.errors.append({
                     "file": filename,
                     "line": tokens[0].line,
                     "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                 })
             else:
                 t_type = current_type
                 if is_array and not t_type.endswith('()'): t_type += "()"
                 scope.define(current_name, t_type, 'Variable')

    def resolve_expression_type(self, tokens, scope, with_stack):
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
            'get', 'put', 'let', 'stop', 'len', 'mid', 'redim', 'preserve', 'erase',
            'friend', 'event', 'implements', 'raiseevent', 'gosub', 'return',
            'lset', 'rset', 'addressof',
            'defbool', 'defbyte', 'defint', 'deflong', 'defcur', 'defsng',
            'defdbl', 'defdec', 'defdate', 'defstr', 'defobj', 'defvar'
        }

        i = 0
        last_resolved_type = None
        last_resolved_kind = None
        last_resolved_name = None
        last_resolved_symbol = None
        expect_member = False
        prev_keyword = None
        
        while i < len(tokens):
            token = tokens[i]
            
            # Check for Implicit Call (Sub style)
            if last_resolved_kind in ('Function', 'Procedure', 'Global') and last_resolved_name and not expect_member:
                 is_arg_start = False
                 if token.type in ('STRING', 'INTEGER', 'FLOAT'): is_arg_start = True
                 elif token.type == 'IDENTIFIER': is_arg_start = True
                 elif token.type == 'OPERATOR' and token.value.lower() in ('-', 'not', 'byval', 'byref'): is_arg_start = True

                 if is_arg_start:
                      arg_tokens = tokens[i:]
                      if report_errors and last_resolved_symbol:
                           self.validate_signature(last_resolved_name, last_resolved_symbol, arg_tokens, filename, token.line, context)

                      args = self.split_args(arg_tokens)
                      for arg in args:
                          self.analyze_statement(arg, scope, filename, context, with_stack, report_errors=report_errors)
                      break

            if token.type == 'IDENTIFIER':
                name = token.value
                last_resolved_name = name
                
                if name.lower() in KEYWORDS and not expect_member:
                    prev_keyword = name.lower()
                    last_resolved_type = None
                    last_resolved_kind = None
                    last_resolved_symbol = None
                    i += 1
                    continue
                
                # Check for Label Definition "Label:" or Named Argument "Arg:="
                if i + 1 < len(tokens) and tokens[i+1].type == 'OPERATOR':
                    if tokens[i+1].value == ':':
                        i += 2
                        prev_keyword = None
                        last_resolved_type = None
                        last_resolved_kind = None
                        continue
                    if tokens[i+1].value == ':=':
                        # Named Argument - skip it and the operator
                        i += 2
                        last_resolved_type = None
                        last_resolved_kind = None
                        continue

                if prev_keyword in ('goto', 'resume', 'gosub'):
                    prev_keyword = None
                    i += 1
                    continue

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
                    last_resolved_kind = 'Unknown' # Member kind resolution not yet implemented
                    last_resolved_symbol = None # Members are not scope symbols
                    expect_member = False
                else:
                    sym = scope.resolve(name)
                    last_resolved_symbol = sym
                    if not sym:
                        # Dynamic ENUM Lookup
                        enum_val = self.resolve_enum(name)
                        if enum_val is not None:
                             last_resolved_type = 'Long'
                             last_resolved_kind = 'EnumItem'
                        else:
                            last_resolved_symbol = None
                            # HEURISTIC: If inside a Form, assume undefined identifier is an implicit Control
                            is_in_form = False
                            curr = scope
                            while curr:
                                if curr.scope_type == 'Form':
                                    is_in_form = True
                                    break
                                curr = curr.parent
                            
                            if is_in_form:
                                last_resolved_type = 'Object'
                                last_resolved_kind = 'Control'
                            else:
                                if report_errors:
                                    self.errors.append({
                                        "file": filename,
                                        "line": token.line,
                                        "message": f"Undefined identifier '{name}' in '{context}'."
                                    })
                                last_resolved_type = 'Unknown'
                                last_resolved_kind = 'Unknown'
                    else:
                        last_resolved_type = sym['type']
                        last_resolved_kind = sym.get('kind', 'Unknown')
                        last_resolved_symbol = sym
                
                prev_keyword = None
                i += 1
            
            elif token.type == 'OPERATOR':
                if token.value == '.':
                    if last_resolved_type is None:
                        if with_stack:
                            last_resolved_type = with_stack[-1]
                            last_resolved_kind = 'Unknown'
                        else:
                            if report_errors:
                                self.errors.append({
                                    "file": filename,
                                    "line": token.line,
                                    "message": f"Invalid or unexpected . reference without With block in '{context}'."
                                })
                            last_resolved_type = 'Unknown'
                            last_resolved_kind = 'Unknown'
                    expect_member = True
                    i += 1
                elif token.value == '(':
                    depth = 1
                    i += 1
                    start_index = i
                    while i < len(tokens) and depth > 0:
                        if tokens[i].type == 'OPERATOR':
                            if tokens[i].value == '(': depth += 1
                            elif tokens[i].value == ')': depth -= 1
                        i += 1
                    
                    end_index = i - 1
                    
                    # Check if we are invoking something that isn't callable
                    # Must be a Variable (not a Procedure/Function) AND have a non-array scalar type
                    if last_resolved_kind == 'Variable' and last_resolved_type in ('String', 'Integer', 'Long', 'Boolean', 'Double', 'Currency', 'Date', 'Single', 'Byte'):
                         if report_errors:
                            self.errors.append({
                                "file": filename,
                                "line": token.line,
                                "message": f"Expected Array or Procedure, got variable '{prev_keyword or 'Unknown'}' of type '{last_resolved_type}'."
                            })

                    # Recursively analyze the content inside the parentheses
                    inner_type = 'Variant'
                    inferred_ret_type = None

                    sub_tokens = []
                    if end_index > start_index:
                        sub_tokens = tokens[start_index : end_index]

                    # Hook for CreateObject("ProgID")
                    if last_resolved_name and last_resolved_name.lower() == 'createobject':
                            if len(sub_tokens) > 0 and sub_tokens[0].type == 'STRING':
                                prog_id = sub_tokens[0].value.strip('"')
                                # Try to resolve ProgID as class
                                if self.config.get_class(prog_id):
                                    inferred_ret_type = prog_id

                    # Hook for Signature Validation
                    if report_errors and last_resolved_symbol:
                            self.validate_signature(last_resolved_name, last_resolved_symbol, sub_tokens, filename, token.line, context)

                    if sub_tokens:
                        inner_type = self.analyze_statement(sub_tokens, scope, filename, context, with_stack, report_errors=report_errors)
                    
                    # Determine result type
                    if last_resolved_type is None:
                        # Grouping (Expression)
                        last_resolved_type = inner_type
                        last_resolved_kind = 'Unknown'
                    else:
                        # Function/Array Call
                        if inferred_ret_type:
                            last_resolved_type = inferred_ret_type
                        elif last_resolved_type.endswith('()'):
                             last_resolved_type = last_resolved_type[:-2]
                        
                        else:
                            # Default Property Logic (e.g. Selection(1) -> Selection.Item(1))
                            # If the type is an object and has an "Item" member, resolve to that type.
                            item_type = self.resolve_member(last_resolved_type, 'Item')
                            if item_type:
                                last_resolved_type = item_type
                            # Otherwise, we keep the original type (it might be a function return type)

                        last_resolved_kind = 'Unknown'
                        
                    expect_member = False
                else:
                    last_resolved_type = None
                    last_resolved_kind = None
                    last_resolved_symbol = None
                    expect_member = False
                    prev_keyword = None
                    i += 1
            
            elif token.type in ('STRING', 'INTEGER', 'FLOAT'):
                last_resolved_type = None
                last_resolved_kind = None
                last_resolved_symbol = None
                expect_member = False
                prev_keyword = None
                i += 1
            else:
                i += 1
        
        return last_resolved_type

    def split_args(self, tokens):
        args = []
        if not tokens: return args

        current_arg = []
        depth = 0
        for t in tokens:
            if t.value == '(':
                depth += 1
                current_arg.append(t)
            elif t.value == ')':
                depth -= 1
                current_arg.append(t)
            elif t.value == ',' and depth == 0:
                args.append(current_arg)
                current_arg = []
            else:
                current_arg.append(t)
        if current_arg or (args and tokens[-1].value == ','):
             args.append(current_arg)
        return args

    def validate_signature(self, name, symbol, arg_tokens, filename, line, context):
        extra = symbol.get('extra')
        if not extra: return

        args = self.split_args(arg_tokens)
        arg_count = len(args)

        min_args = 0
        max_args = 999

        if isinstance(extra, ProcedureNode):
             min_args = 0
             max_args = len(extra.args)
             has_param_array = False

             for arg in extra.args:
                  if getattr(arg, 'is_paramarray', False):
                       has_param_array = True

                  if not arg.is_optional and not getattr(arg, 'is_paramarray', False):
                       min_args += 1

             if has_param_array:
                  max_args = 999
        elif isinstance(extra, dict):
             min_args = extra.get('min_args', 0)
             max_args = extra.get('max_args', 999)

        if arg_count < min_args:
             self.errors.append({
                 "file": filename,
                 "line": line,
                 "message": f"Argument count mismatch for '{name}': Expected at least {min_args}, got {arg_count}."
             })
        if arg_count > max_args:
             self.errors.append({
                 "file": filename,
                 "line": line,
                 "message": f"Argument count mismatch for '{name}': Expected at most {max_args}, got {arg_count}."
             })

    def resolve_member(self, type_name, member_name):
        return self._resolve_member_internal(type_name, member_name)

    def _resolve_member_internal(self, type_name, member_name):
        # Handle Qualified Types
        if '.' in type_name:
             simple_name = type_name.split('.')[-1]
             res = self._resolve_member_base(type_name, member_name)
             if res: return res
             res = self._resolve_member_base(simple_name, member_name)
             if res: return res
             return None
        
        return self._resolve_member_base(type_name, member_name)

    def _resolve_member_base(self, type_name, member_name):
        # 0. Check UDTs
        if type_name.lower() in self.udts:
            udt = self.udts[type_name.lower()]
            for m in udt.members:
                if m.name.lower() == member_name.lower():
                    return m.type_name
        
        # 1. Check Config Classes (Loaded from Model)
        cls_def = self.config.get_class(type_name)
        if cls_def:
            members = cls_def.get("members", {})
            for m_name, m_def in members.items():
                if m_name.lower() == member_name.lower():
                    return m_def.get("type", "Variant")
        
        # 2. Check Standard Modules
        for mod in self.modules:
            if mod.module_type == 'Module' and mod.name.lower() == type_name.lower():
                for v in mod.variables:
                    if v.name.lower() == member_name.lower() and v.scope.lower() in ('public', 'global', 'friend'):
                        return v.type_name
                for p in mod.procedures:
                    if p.name.lower() == member_name.lower() and p.scope.lower() in ('public', 'friend'):
                         return p.return_type
        
        # 3. Check Project Classes
        for mod in self.modules:
            if mod.module_type in ('Class', 'Form') and mod.name.lower() == type_name.lower():
                 for v in mod.variables:
                     if v.name.lower() == member_name.lower() and v.scope.lower() in ('public', 'global', 'friend'):
                         return v.type_name
                 for p in mod.procedures:
                     if p.name.lower() == member_name.lower() and p.scope.lower() in ('public', 'friend'):
                         return p.return_type
                 
                 # FALLBACK: If it's a Form, check the 'UserForm' class definition (from config)
                 if mod.module_type == 'Form':
                     userform_cls = self.config.get_class('UserForm')
                     if userform_cls:
                         members = userform_cls.get('members', {})
                         for m_name, m_def in members.items():
                             if m_name.lower() == member_name.lower():
                                 return m_def.get('type', 'Variant')
                     
                     # FALLBACK 2: Implicit Controls (e.g. Me.txtBox)
                     # Since we can't parse controls from .frm text, assume any other member is a control
                     return 'Object'
                 
                 # FALLBACK: ThisDocument (Special project class that acts as a Document)
                 if mod.name.lower() == 'thisdocument':
                     doc_cls = self.config.get_class('Document') or self.config.get_class('IVDocument')
                     if doc_cls:
                         members = doc_cls.get('members', {})
                         for m_name, m_def in members.items():
                             if m_name.lower() == member_name.lower():
                                 return m_def.get('type', 'Variant')

        return None

    def is_label(self, tokens):
        if len(tokens) >= 2:
            if tokens[0].type == 'IDENTIFIER' and tokens[1].type == 'OPERATOR' and tokens[1].value == ':':
                return True
        return False

    def is_ignorable(self, tokens):
        if not tokens: return True
        for t in tokens:
            if t.type != 'COMMENT':
                return False
        return True

    def is_unconditional_jump(self, tokens):
        if not tokens: return False
        t0 = tokens[0]
        val = t0.value.lower()

        if val == 'goto':
            # Check if strictly GoTo Label (Simple GoTo is 2 tokens, plus maybe a colon if parsed that way)
            if len(tokens) >= 2 and tokens[1].type == 'IDENTIFIER':
                 # Ensure it's not On Error GoTo ...
                 if len(tokens) == 2: return True
                 if len(tokens) == 3 and tokens[2].value == ':': return True

        if val == 'exit':
            if len(tokens) >= 2:
                t1 = tokens[1].value.lower()
                if t1 in ('sub', 'function', 'property'):
                    return True

        if val == 'end':
            if len(tokens) == 1: return True
            if len(tokens) == 2 and tokens[1].value == ':': return True

        return False

    def is_control_flow_boundary(self, tokens):
        if not tokens: return False
        val = tokens[0].value.lower()

        if val in ('else', 'elseif', 'next', 'loop', 'wend', 'case'):
            return True

        if val == 'end':
            if len(tokens) >= 2:
                 val2 = tokens[1].value.lower()
                 if val2 in ('if', 'select', 'with'):
                     return True

        return False

    def resolve_enum(self, name):
        # Look up enum constants
        enums = self.config.object_model.get("enums", {})
        for enum_name, members in enums.items():
            if name.lower() in [m.lower() for m in members.keys()]:
                return members.get(name) # Or handle case insensitive lookup properly
            # Case insensitive check
            for m_key, m_val in members.items():
                if m_key.lower() == name.lower():
                    return m_val
        return None
