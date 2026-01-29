from .lexer import Lexer, Token

class Node:
    pass

class VariableNode(Node):
    def __init__(self, name, type_name, scope='Private', is_optional=False, is_paramarray=False, mechanism='ByRef'):
        self.name = name
        self.type_name = type_name
        self.scope = scope # Dim (Local), Private, Public, Global
        self.is_optional = is_optional
        self.is_paramarray = is_paramarray
        self.mechanism = mechanism

    def __repr__(self):
        return f"Var({self.name} As {self.type_name} [{self.mechanism}])"

class StatementNode(Node):
    def __init__(self, tokens):
        self.tokens = tokens
    def __repr__(self):
        return f"Stmt({len(self.tokens)} tokens)"

class WithNode(Node):
    def __init__(self, expr_tokens, body):
        self.expr_tokens = expr_tokens
        self.body = body # List of StatementNode or WithNode
    def __repr__(self):
        return f"With(expr, {len(self.body)} stmts)"

class ProcedureNode(Node):
    def __init__(self, name, proc_type, return_type='Variant', scope='Public', is_declare=False, lib_name=None, alias_name=None):
        self.name = name
        self.proc_type = proc_type # Sub, Function, Property Get/Set/Let
        self.return_type = return_type
        self.scope = scope
        self.is_declare = is_declare
        self.lib_name = lib_name
        self.alias_name = alias_name
        self.args = [] # List of VariableNode
        self.locals = [] # List of VariableNode
        self.body = [] # List of nodes (StatementNode, WithNode)

    def __repr__(self):
        decl = "Declare " if self.is_declare else ""
        return f"{decl}{self.proc_type} {self.name}() As {self.return_type}"

class TypeNode(Node):
    def __init__(self, name, scope='Public'):
        self.name = name
        self.scope = scope
        self.members = [] # List of VariableNode

    def __repr__(self):
        return f"Type {self.name} ({len(self.members)} members)"

class ModuleNode(Node):
    def __init__(self, filename, module_type='Module'):
        self.filename = filename
        self.name = "Unknown"
        self.module_type = module_type # Module, Class, Form
        self.attributes = {}
        self.variables = [] # Module-level variables
        self.procedures = []
        self.types = {} # User Defined Types

class IfNode(Node):
    def __init__(self, condition_tokens, true_block, else_blocks=None, else_block=None):
        self.condition_tokens = condition_tokens
        self.true_block = true_block
        self.else_blocks = else_blocks if else_blocks else [] # List of (condition_tokens, block)
        self.else_block = else_block

class FormParser:
    """Parses the GUI definition block of .frm files."""
    def parse(self, content):
        controls = []
        import re
        begin_pat = re.compile(r'^\s*Begin\s+([\w\.]+)\s+(\w+)', re.MULTILINE)
        
        for match in begin_pat.finditer(content):
            cls_type = match.group(1)
            name = match.group(2)
            if '.' in cls_type:
                cls_type = cls_type.split('.')[-1]
            controls.append(VariableNode(name, cls_type, scope='Public'))
            
        return controls

class VBAParser:
    def __init__(self, tokens, filename="Unknown"):
        self.tokens = tokens
        self.filename = filename
        self.pos = 0
        self.current_token = None
        self.advance()

    def advance(self):
        if self.pos < len(self.tokens):
            self.current_token = self.tokens[self.pos]
            self.pos += 1
        else:
            self.current_token = Token('EOF', '', -1, -1)

    def peek(self):
        if self.pos < len(self.tokens):
            return self.tokens[self.pos]
        return Token('EOF', '', -1, -1)

    def consume(self, type_name=None, value=None):
        if type_name and self.current_token.type != type_name:
            return False
        if value and self.current_token.value.lower() != value.lower():
            return False
        self.advance()
        return True

    def match(self, type_name=None, value=None):
        if type_name and self.current_token.type != type_name:
            return False
        if value and self.current_token.value.lower() != value.lower():
            return False
        return True

    def parse_module(self):
        module = ModuleNode("Unknown")
        
        while self.current_token.type != 'EOF':
            if self.match('IDENTIFIER', 'Attribute'):
                self.parse_attribute(module)
            elif self.match('IDENTIFIER', 'Option'):
                self.consume()
                self.consume_statement()
            elif self.match('IDENTIFIER', 'Implements'):
                self.consume()
                self.consume_statement()
            elif self.current_token.type == 'IDENTIFIER' and self.current_token.value.lower().startswith('def'):
                # Handle DefInt, DefBool, etc.
                self.consume()
                self.consume_statement()
            elif self.match('IDENTIFIER', 'Public') or self.match('IDENTIFIER', 'Private') or self.match('IDENTIFIER', 'Friend') or self.match('IDENTIFIER', 'Dim') or self.match('IDENTIFIER', 'Const') or self.match('IDENTIFIER', 'Global'):
                self.parse_declaration(module)
            elif self.match('IDENTIFIER', 'Sub') or self.match('IDENTIFIER', 'Function') or self.match('IDENTIFIER', 'Property'):
                self.procedures_parse(module, 'Public') 
            elif self.match('IDENTIFIER', 'Type'):
                self.parse_udt(module)
            elif self.match('IDENTIFIER', 'Event'):
                # Handle implicit public Event
                # Reuse parse_declaration logic or duplicate event parsing?
                # parse_declaration expects current_token to be scope.
                # Here we already consumed 'Event' if we use match? No match doesn't consume if successful?
                # match returns True but does NOT consume unless I call advance?
                # Wait, match method:
                # if type_name and self.current_token.type != type_name: return False
                # return True (Does NOT advance)

                # So if I match Event, current token is Event.
                # I can manually parse it here.
                self.consume() # Event
                event_name = "Unknown"
                if self.current_token.type == 'IDENTIFIER':
                    event_name = self.current_token.value
                    self.advance()

                proc = ProcedureNode(event_name, 'Event', scope='Public')
                if self.match('OPERATOR', '('):
                    self.parse_arg_list(proc)
                self.consume_statement()
                module.procedures.append(proc)

            elif self.match('IDENTIFIER', 'Enum'):
                self.parse_enum(module, 'Public')
            elif self.match('NEWLINE'):
                self.advance()
            else:
                self.consume_statement()
        
        return module

    def consume_statement(self):
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            if self.current_token.type == 'OPERATOR' and self.current_token.value == ':':
                 break
            self.advance()
        if self.current_token.type == 'NEWLINE':
            self.advance()

    def parse_attribute(self, module):
        self.consume('IDENTIFIER', 'Attribute')
        
        attr_name = "Unknown"
        if self.current_token.type == 'IDENTIFIER':
            attr_name = self.current_token.value
            self.advance()
            
        self.consume('OPERATOR', '=')
        
        attr_value = "Unknown"
        if self.current_token.type == 'STRING':
            attr_value = self.current_token.value.strip('"')
            self.advance()
        elif self.current_token.type == 'IDENTIFIER': # True/False
            attr_value = self.current_token.value
            self.advance()
        
        module.attributes[attr_name] = attr_value
        
        if attr_name == 'VB_Name':
            module.name = attr_value
            
        self.consume_statement()

    def parse_declaration(self, module):
        scope = self.current_token.value # Public, Private, Dim
        self.advance()
        
        # Handle Event
        # [Public|Private|Friend] Event Name(...)
        if self.match('IDENTIFIER', 'Event'):
            self.advance()
            event_name = "Unknown"
            if self.current_token.type == 'IDENTIFIER':
                event_name = self.current_token.value
                self.advance()

            proc = ProcedureNode(event_name, 'Event', scope=scope)

            if self.match('OPERATOR', '('):
                self.parse_arg_list(proc)

            self.consume_statement()
            module.procedures.append(proc)
            return

        # Handle Declare
        # [Public|Private] Declare [PtrSafe] Sub/Function ...
        if self.match('IDENTIFIER', 'Declare'):
            self.advance() # consume Declare
            
            # Optional PtrSafe
            is_ptrsafe = False
            if self.match('IDENTIFIER', 'PtrSafe'):
                self.advance()
                is_ptrsafe = True
            
            proc_type = "Sub"
            if self.match('IDENTIFIER', 'Function'):
                proc_type = "Function"
                self.advance()
            elif self.match('IDENTIFIER', 'Sub'):
                self.advance()
            
            proc_name = "Unknown"
            if self.current_token.type == 'IDENTIFIER':
                proc_name = self.current_token.value
                self.advance()
            
            # Lib "..."
            lib_name = None
            if self.match('IDENTIFIER', 'Lib'):
                self.advance()
                if self.current_token.type == 'STRING':
                    lib_name = self.current_token.value
                    self.advance()
            
            # Alias "..."
            alias_name = None
            if self.match('IDENTIFIER', 'Alias'):
                self.advance()
                if self.current_token.type == 'STRING':
                    alias_name = self.current_token.value
                    self.advance()
            
            proc = ProcedureNode(proc_name, proc_type, scope=scope, is_declare=True, lib_name=lib_name, alias_name=alias_name)
            
            # Args (...)
            if self.match('OPERATOR', '('):
                self.parse_arg_list(proc)
            
            # Return type
            if self.match('IDENTIFIER', 'As'):
                self.advance()
                proc.return_type = self.parse_type_signature()
                
            self.consume_statement()
            module.procedures.append(proc)
            return

        if self.match('IDENTIFIER', 'Sub') or self.match('IDENTIFIER', 'Function') or self.match('IDENTIFIER', 'Property'):
            self.procedures_parse(module, scope)
            return

        # Handle 'Type' (Public Type ...)
        if self.match('IDENTIFIER', 'Type'):
            self.parse_udt(module, scope=scope)
            return

        # Handle 'Enum' (Public Enum ...)
        if self.match('IDENTIFIER', 'Enum'):
            self.parse_enum(module, scope=scope)
            return

        # Check if Const
        if scope.lower() in ('public', 'private', 'global', 'friend'):
             if self.match('IDENTIFIER', 'Const'):
                 self.advance()
        
        # Check if WithEvents
        if self.match('IDENTIFIER', 'WithEvents'):
            self.advance()

        # Dim x As Type
        while True:
            if self.current_token.type == 'IDENTIFIER':
                var_name = self.current_token.value
                self.advance()
                var_type = 'Variant'
                
                if self.match('IDENTIFIER', 'As'):
                    self.advance()
                    var_type = self.parse_type_signature()
                
                # Handle array: x(10)
                if self.match('OPERATOR', '('):
                    while self.current_token.type != 'EOF' and not self.match('OPERATOR', ')'):
                        self.advance()
                    self.consume('OPERATOR', ')')
                    var_type += "()" 

                # Handle initialization = ...
                if self.match('OPERATOR', '='):
                     while self.current_token.type not in ('NEWLINE', 'EOF') and not self.match('OPERATOR', ','):
                         self.advance()

                module.variables.append(VariableNode(var_name, var_type, scope))
            
            if self.match('OPERATOR', ','):
                self.advance()
                continue
            else:
                break
        
        self.consume_statement()

    def parse_type_signature(self):
        # Ignore 'New' keyword if present
        if self.match('IDENTIFIER', 'New'):
            self.advance()

        type_parts = []
        while self.current_token.type == 'IDENTIFIER':
            type_parts.append(self.current_token.value)
            self.advance()
            if self.match('OPERATOR', '.'):
                self.advance()
                type_parts.append('.')
            else:
                break
        return "".join(type_parts)

    def procedures_parse(self, module, scope):
        proc_type = self.current_token.value 
        self.advance()
        
        if self.match('IDENTIFIER', 'Get') or self.match('IDENTIFIER', 'Let') or self.match('IDENTIFIER', 'Set'):
            proc_type += " " + self.current_token.value
            self.advance()
            
        proc_name = "Unknown"
        if self.current_token.type == 'IDENTIFIER':
            proc_name = self.current_token.value
            self.advance()
            
        proc = ProcedureNode(proc_name, proc_type, scope=scope)
        
        # Args
        if self.match('OPERATOR', '('):
            self.parse_arg_list(proc)
            
        if self.match('IDENTIFIER', 'As'):
            self.advance()
            proc.return_type = self.parse_type_signature()
            
        self.consume_statement()
        
        # Parse Body Block
        end_marker = proc_type.split()[0].lower() # Sub, Function, Property
        proc.body = self.parse_block(end_markers=[f"End {end_marker}", "End"])
        
        # Ensure we consumed End Sub/Function
        if self.match('IDENTIFIER', 'End'):
             self.advance()
             if self.current_token.value.lower() == end_marker:
                 self.advance()
        self.consume_statement()

        module.procedures.append(proc)

    def parse_block(self, end_markers):
        """Recursively parses statements until an end marker is found."""
        nodes = []
        
        while self.current_token.type != 'EOF':
            # Check for End Markers
            if self.current_token.type == 'IDENTIFIER' and self.current_token.value.lower() == 'end':
                # Peek to see what kind of End it is
                peek = self.peek()
                combined = f"End {peek.value}".lower()
                
                # Check if it matches an expected marker
                for marker in end_markers:
                    if marker.lower() == combined:
                        return nodes
                    if marker.lower() == 'end': # Naked End (e.g. End Select vs End) - unlikely for blocks except maybe Sub
                         pass

            # Also check for intermediate markers (Else, ElseIf, Loop, Next) provided in end_markers
            # For "Next", it might be "Next i", so we need to be careful.
            if self.current_token.type == 'IDENTIFIER':
                val = self.current_token.value.lower()
                # Check directly if the current token matches a marker (e.g. "Loop", "Next", "Else")
                for marker in end_markers:
                    if marker.lower().split()[0] == val:
                         # Potential match. 
                         # If marker is "Next", and we have "Next i", it's a match.
                         # If marker is "Else", and we have "Else", it's a match.
                         return nodes

                # VALIDATION: Check for unexpected block terminators
                if val in ('next', 'loop', 'else', 'elseif', 'wend'):
                    # Found a block keyword that was NOT in end_markers -> Unexpected
                    print(f"{self.filename}:{self.current_token.line}: Syntax Error: Unexpected '{self.current_token.value}'")
                    # We consume it to avoid infinite loop, but it's an error
                    self.consume_statement()
                    continue

                if val == 'end':
                    peek_val = self.peek().value.lower()
                    if peek_val in ('if', 'select', 'with', 'function', 'sub', 'property'):
                        # Found End X that was NOT in end_markers -> Unexpected
                        print(f"{self.filename}:{self.current_token.line}: Syntax Error: Unexpected 'End {self.peek().value}'")
                        self.advance() # End
                        self.advance() # X
                        self.consume_statement()
                        continue

            # Parse Statements
            if self.match('IDENTIFIER', 'With'):
                nodes.append(self.parse_with())
            
            elif self.match('IDENTIFIER', 'If'):
                stmt = self.parse_if_stmt()
                if stmt: nodes.append(stmt)
            
            elif self.match('IDENTIFIER', 'For'):
                nodes.append(self.parse_for())
            
            elif self.match('IDENTIFIER', 'Do'):
                nodes.append(self.parse_do())
                
            elif self.match('IDENTIFIER', 'Select'):
                nodes.append(self.parse_select())

            elif self.match('IDENTIFIER', 'While'):
                 nodes.append(self.parse_while())

            elif self.match('IDENTIFIER', 'Dim') or self.match('IDENTIFIER', 'Static'):
                stmt = self.collect_statement()
                nodes.append(StatementNode(stmt))
                
            else:
                # Normal Statement
                stmt = self.collect_statement()
                if stmt:
                    nodes.append(StatementNode(stmt))
                else:
                    if self.current_token.type == 'NEWLINE':
                        self.advance()

        return nodes

    def parse_while(self):
        self.consume('IDENTIFIER', 'While')
        condition_tokens = self.collect_statement()  # Everything until newline
        
        body = self.parse_block(end_markers=["Wend"])
        
        self.consume('IDENTIFIER', 'Wend')
        self.consume_statement()
        
        return StatementNode(condition_tokens) # Placeholder for WhileNode

    def parse_with(self):
        self.consume('IDENTIFIER', 'With')
        expr_tokens = []
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            expr_tokens.append(self.current_token)
            self.advance()
        self.consume_statement()
        
        body = self.parse_block(end_markers=["End With"])
        
        self.consume('IDENTIFIER', 'End')
        self.consume('IDENTIFIER', 'With')
        self.consume_statement()
        
        return WithNode(expr_tokens, body)

    def parse_if_stmt(self):
        # If <condition> Then <newline> [Block]
        # If <condition> Then <statement> [Else <statement>] [newline] [Single Line]
        
        self.consume('IDENTIFIER', 'If')
        
        # Scavenge tokens until 'Then'
        condition_tokens = []
        while self.current_token.type != 'EOF':
            if self.match('IDENTIFIER', 'Then'):
                break
            condition_tokens.append(self.current_token)
            self.advance()
            
        if not self.match('IDENTIFIER', 'Then'):
             # Syntax Error: Missing Then
             print(f"Syntax Error: Missing 'Then' at line {self.current_token.line}")
             self.consume_statement() # Recover
             return None
             
        self.consume('IDENTIFIER', 'Then')
        
        # Check for single line vs block
        if self.current_token.type == 'NEWLINE' or self.current_token.type == 'COMMENT':
             # Block If
             self.consume_statement()
             
             # Parse True Block
             # We stop at Else, ElseIf, or End If
             true_block = self.parse_block(end_markers=["Else", "ElseIf", "End If"])
             
             else_blocks = []
             else_block = None
             
             while True:
                 tok = self.current_token
                 if tok.type == 'IDENTIFIER':
                     val = tok.value.lower()
                     
                     if val == 'elseif':
                         self.advance()
                         # Parse condition Then
                         elseif_cond = []
                         while not self.match('IDENTIFIER', 'Then') and self.current_token.type != 'EOF':
                             elseif_cond.append(self.current_token)
                             self.advance()
                         self.consume('IDENTIFIER', 'Then')
                         self.consume_statement()
                         
                         block = self.parse_block(end_markers=["Else", "ElseIf", "End If"])
                         else_blocks.append((elseif_cond, block))
                     
                     elif val == 'else':
                         self.advance()
                         self.consume_statement()
                         else_block = self.parse_block(end_markers=["End If"])
                         # Do not break here. Let loop consume End If.
                         pass
                     
                     elif val == 'end':
                         peek = self.peek()
                         if peek.value.lower() == 'if':
                             self.advance() # End
                             self.advance() # If
                             self.consume_statement()
                         break
                     
                     else:
                         break
                 else:
                     break
             
             return IfNode(condition_tokens, true_block, else_blocks, else_block)
             
        else:
             # Single Line If
             # Parse until Newline
             # We must consume ALL statements on this line, including those separated by colons.
             condition_tokens.extend(self.collect_statement(consume_newline=False))
             
             while self.current_token.type != 'NEWLINE' and self.current_token.type != 'EOF':
                 # Loop for chained statements `If ... Then ... : ...`
                 
                 # The previous collect_statement(consume_newline=False) either stopped at ':' (consumed) or NEWLINE (not consumed).
                 # If it stopped at ':', self.current_token is the next statement's start.
                 # If it stopped at NEWLINE, we wouldn't be in this loop (due to while condition).
                 
                 # So we simply collect the next statement.
                 stmt_tokens = self.collect_statement(consume_newline=False)
                 condition_tokens.extend(stmt_tokens)
             
             # Consume the final newline
             if self.match('NEWLINE'):
                 self.advance()



    def parse_for(self):
        self.consume('IDENTIFIER', 'For')
        # ... = ... To ...
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            self.advance()
        self.consume_statement()
        
        body = self.parse_block(end_markers=["Next"])
        
        self.consume('IDENTIFIER', 'Next')
        # Optional variable
        if self.current_token.type == 'IDENTIFIER':
            self.advance()
        self.consume_statement()
        
        return StatementNode([]) # Placeholder

    def parse_do(self):
        self.consume('IDENTIFIER', 'Do')
        while self.current_token.type not in ('NEWLINE', 'EOF'):
             self.advance()
        self.consume_statement()
        
        body = self.parse_block(end_markers=["Loop"])
        
        self.consume('IDENTIFIER', 'Loop')
        while self.current_token.type not in ('NEWLINE', 'EOF'):
             self.advance()
        self.consume_statement()
        
        return StatementNode([])

    def parse_select(self):
        self.consume('IDENTIFIER', 'Select')
        self.consume('IDENTIFIER', 'Case')
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            self.advance()
        self.consume_statement()
        
        # Select block ends with End Select
        # Inside we have Case ...
        # But parse_block effectively handles the content. 
        # We just need to stop at End Select.
        # But wait, 'Case' is a separator?
        # A simple block parse works if we consider Case as statements inside the block?
        # OR we can treat it like If/ElseIf.
        # For simple nesting check, just parsing until End Select is enough, 
        # as long as we don't error on Case statements. 
        # Case statements will be picked up as normal statements.
        
        body = self.parse_block(end_markers=["End Select"])
        
        self.consume('IDENTIFIER', 'End')
        self.consume('IDENTIFIER', 'Select')
        self.consume_statement()
        
        return StatementNode([])

    def collect_statement(self, consume_newline=True):
        tokens = []
        while self.current_token.type != 'NEWLINE' and self.current_token.type != 'EOF':
             tokens.append(self.current_token)
             
             # Check for statement separator ':'
             if self.current_token.type == 'OPERATOR' and self.current_token.value == ':':
                 self.advance()
                 # Break statement here, but include the colon so analyzer can detect labels "Label:"
                 return tokens

             self.advance()
        if consume_newline and self.current_token.type == 'NEWLINE':
            self.advance()
        return tokens

    def parse_arg_list(self, proc):
        self.consume('OPERATOR', '(')
        while not self.match('OPERATOR', ')') and self.current_token.type != 'EOF':
            is_optional = False
            is_paramarray = False
            mechanism = 'ByRef'

            while self.match('IDENTIFIER', 'Optional') or self.match('IDENTIFIER', 'ByVal') or self.match('IDENTIFIER', 'ByRef') or self.match('IDENTIFIER', 'ParamArray'):
                val = self.current_token.value.lower()
                if val == 'optional': is_optional = True
                if val == 'paramarray':
                    is_paramarray = True
                    mechanism = 'ParamArray'
                if val == 'byval': mechanism = 'ByVal'
                if val == 'byref': mechanism = 'ByRef'
                self.advance()
            
            if self.current_token.type == 'IDENTIFIER':
                arg_name = self.current_token.value
                self.advance()

                is_array = False
                # Check for array parens on name: arr()
                if self.match('OPERATOR', '('):
                        self.advance()
                        self.consume('OPERATOR', ')')
                        is_array = True

                arg_type = 'Variant'
                if self.match('IDENTIFIER', 'As'):
                    self.advance()
                    arg_type = self.parse_type_signature()
                
                # Check for array parens on type (rare but supported by my parser previously)
                if self.match('OPERATOR', '('):
                        self.advance()
                        self.consume('OPERATOR', ')')
                        is_array = True

                if is_array and not arg_type.endswith('()'):
                     arg_type += "()"

                # Handle Default Value (= ...)
                if self.match('OPERATOR', '='):
                    self.advance()
                    # Skip until ',' or ')'
                    while self.current_token.type != 'EOF':
                         if self.current_token.type == 'OPERATOR' and self.current_token.value in (',', ')'):
                             break
                         self.advance()

                proc.args.append(VariableNode(arg_name, arg_type, 'Local', is_optional=is_optional, is_paramarray=is_paramarray, mechanism=mechanism))
            
            if self.match('OPERATOR', ','):
                self.advance()
            elif self.current_token.type != 'EOF' and not self.match('OPERATOR', ')'):
                    self.advance()
        self.consume('OPERATOR', ')')

    def parse_udt(self, module, scope='Public'):
        # NOTE: Caller (parse_module) consumes 'Type' BEFORE calling this? No.
        # "elif self.match('IDENTIFIER', 'Type'): self.parse_udt(module)"
        # But wait, inside parse_module, it checks current_token 'Type'.
        # Inside parse_declaration, it checks current_token 'Type'.
        # self.consume('IDENTIFIER', 'Type') needs to succeed.
        
        self.consume('IDENTIFIER', 'Type')
        type_name = self.current_token.value
        self.advance()
        self.consume_statement()
        
        udt = TypeNode(type_name, scope)
        
        while self.current_token.type != 'EOF':
            # Check for End Type
            if self.match('IDENTIFIER', 'End') and self.peek().value.lower() == 'type':
                self.advance() # End
                self.advance() # Type
                self.consume_statement()
                break
            
            # Parse Member: Name As Type
            if self.current_token.type == 'IDENTIFIER':
                var_name = self.current_token.value
                self.advance()
                
                var_type = 'Variant'
                if self.match('IDENTIFIER', 'As'):
                    self.advance()
                    var_type = self.parse_type_signature()
                
                # Check for array
                if self.match('OPERATOR', '('):
                    while not self.match('OPERATOR', ')') and self.current_token.type != 'EOF':
                        self.advance()
                    self.consume('OPERATOR', ')')
                    var_type += "()"
                
                # Check for * N (Fixed length string) - simplified ignore
                if self.match('OPERATOR', '*'):
                    self.advance()
                    self.advance() # length
                
                udt.members.append(VariableNode(var_name, var_type, 'Public'))
            
            self.consume_statement()
            
        module.types[type_name] = udt

    def parse_enum(self, module, scope='Public'):
        self.consume('IDENTIFIER', 'Enum')
        enum_name = self.current_token.value
        self.advance()
        self.consume_statement()

        # Enums are basically Longs with named constants
        # We need to register the Enum Type AND the Enum Members as global/module constants

        # Create a TypeNode to represent the Enum type itself?
        # Or just store members?
        # Analyzer needs to know EnumName is a valid Type.
        udt = TypeNode(enum_name, scope) # Reuse TypeNode for simplicity

        while self.current_token.type != 'EOF':
            if self.match('IDENTIFIER', 'End') and self.peek().value.lower() == 'enum':
                self.advance()
                self.advance()
                self.consume_statement()
                break

            # Member: Name = Value
            if self.current_token.type == 'IDENTIFIER':
                member_name = self.current_token.value
                self.advance()

                # Enum members are constants.
                # We should register them in the module's constants/variables list?
                # Or a specific Enum list?
                # Analyzer expects module.types for types.
                # For members, it checks variables/constants?

                # Let's treat them as Public Constants for now.
                # But we also want to support `Dim x As EnumName`.

                # So we register the Enum Type in module.types
                # AND we register the members as module-level variables (Consts)

                var = VariableNode(member_name, 'Long', scope) # Enum members are Long
                module.variables.append(var)
                udt.members.append(var)

                if self.match('OPERATOR', '='):
                    self.advance()
                    # Skip value
                    while self.current_token.type not in ('NEWLINE', 'EOF', 'COMMENT'):
                        self.advance()

            self.consume_statement()

        module.types[enum_name] = udt
