from .lexer import Lexer, Token

class Node:
    pass

class VariableNode(Node):
    def __init__(self, name, type_name, scope='Private'):
        self.name = name
        self.type_name = type_name
        self.scope = scope # Dim (Local), Private, Public, Global

    def __repr__(self):
        return f"Var({self.name} As {self.type_name})"

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
    def __init__(self, name, proc_type, return_type='Variant', scope='Public'):
        self.name = name
        self.proc_type = proc_type # Sub, Function, Property Get/Set/Let
        self.return_type = return_type
        self.scope = scope
        self.args = [] # List of VariableNode
        self.locals = [] # List of VariableNode
        self.body = [] # List of nodes (StatementNode, WithNode)

    def __repr__(self):
        return f"{self.proc_type} {self.name}() As {self.return_type}"

class ModuleNode(Node):
    def __init__(self, filename, module_type='Module'):
        self.filename = filename
        self.name = "Unknown"
        self.module_type = module_type # Module, Class, Form
        self.attributes = {}
        self.variables = [] # Module-level variables
        self.procedures = []
        self.types = {} # User Defined Types

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
    def __init__(self, tokens):
        self.tokens = tokens
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
            elif self.match('IDENTIFIER', 'Public') or self.match('IDENTIFIER', 'Private') or self.match('IDENTIFIER', 'Dim') or self.match('IDENTIFIER', 'Const') or self.match('IDENTIFIER', 'Global'):
                self.parse_declaration(module)
            elif self.match('IDENTIFIER', 'Sub') or self.match('IDENTIFIER', 'Function') or self.match('IDENTIFIER', 'Property'):
                self.procedures_parse(module, 'Public') 
            elif self.match('IDENTIFIER', 'Type'):
                self.parse_udt(module)
            elif self.match('NEWLINE'):
                self.advance()
            else:
                self.consume_statement()
        
        return module

    def consume_statement(self):
        while self.current_token.type != 'NEWLINE' and self.current_token.type != 'EOF':
            self.advance()
        if self.current_token.type == 'NEWLINE':
            self.advance()

    def parse_attribute(self, module):
        self.consume('IDENTIFIER', 'Attribute')
        self.consume('IDENTIFIER', 'VB_Name')
        self.consume('OPERATOR', '=')
        if self.current_token.type == 'STRING':
            module.name = self.current_token.value.strip('"')
            self.advance()
        self.consume_statement()

    def parse_declaration(self, module):
        scope = self.current_token.value # Public, Private, Dim
        self.advance()
        
        if self.match('IDENTIFIER', 'Sub') or self.match('IDENTIFIER', 'Function') or self.match('IDENTIFIER', 'Property'):
            self.procedures_parse(module, scope)
            return

        # Check if Const
        if scope.lower() in ('public', 'private', 'global'):
             if self.match('IDENTIFIER', 'Const'):
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
            self.advance()
            while not self.match('OPERATOR', ')') and self.current_token.type != 'EOF':
                while self.match('IDENTIFIER', 'Optional') or self.match('IDENTIFIER', 'ByVal') or self.match('IDENTIFIER', 'ByRef') or self.match('IDENTIFIER', 'ParamArray'):
                    self.advance()
                
                if self.current_token.type == 'IDENTIFIER':
                    arg_name = self.current_token.value
                    self.advance()
                    arg_type = 'Variant'
                    if self.match('IDENTIFIER', 'As'):
                        self.advance()
                        arg_type = self.parse_type_signature()
                    
                    if self.match('OPERATOR', '('):
                         self.advance()
                         self.consume('OPERATOR', ')')
                         arg_type += "()"
                         
                    proc.args.append(VariableNode(arg_name, arg_type, 'Local'))
                
                if self.match('OPERATOR', ','):
                    self.advance()
                elif self.current_token.type != 'EOF' and not self.match('OPERATOR', ')'):
                     self.advance()
            self.consume('OPERATOR', ')')
            
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
            # Helper to check if current token sequence matches "End Sub", "End With", etc.
            is_end = False
            
            # Simple check: If current is 'End' and next is X, check "End X"
            # Or if marker is just "End"
            if self.current_token.type == 'IDENTIFIER':
                # Check strict multi-word markers first
                # We need to peek ahead without consuming
                pass 

            # Let's peek
            if self.current_token.value.lower() == 'end':
                peek = self.peek()
                combined = f"End {peek.value}".lower()
                
                # Check matches
                found_match = False
                for marker in end_markers:
                    if marker.lower() == combined:
                        return nodes # Stop parsing block, don't consume markers here
                    if marker.lower() == 'end': # Naked End
                         # But wait, End Sub shouldn't match End if End Sub is expected
                         pass
                
                # Specific logic:
                # If we expect "End With", and we see "End With", return.
                # If we expect "End Sub", and we see "End Sub", return.
            
            # Handling "End With" vs "End Sub"
            if self.match('IDENTIFIER', 'End'):
                peek = self.current_token # because match advanced? No, match does NOT advance if false. But check manual.
                # match returns True/False.
                # Actually, I need to check WITHOUT consuming to know if I should stop.
                
                # Re-implement robust check
                current_val = self.current_token.value.lower()
                next_val = self.peek().value.lower()
                
                if current_val == 'end':
                     combined = f"end {next_val}"
                     # Check if combined matches any marker
                     for m in end_markers:
                         if m.lower() == combined:
                             return nodes
            
            # Parse Statements
            if self.match('IDENTIFIER', 'With'):
                # With Block
                self.consume('IDENTIFIER', 'With')
                expr_tokens = []
                while self.current_token.type not in ('NEWLINE', 'EOF'):
                    expr_tokens.append(self.current_token)
                    self.advance()
                self.consume_statement()
                
                body = self.parse_block(end_markers=["End With"])
                nodes.append(WithNode(expr_tokens, body))
                
                # Consume End With
                if self.match('IDENTIFIER', 'End') and self.peek().value.lower() == 'with':
                    self.advance() # End
                    self.advance() # With
                    self.consume_statement()
            
            elif self.match('IDENTIFIER', 'Dim') or self.match('IDENTIFIER', 'Static'):
                # Local Decl - parse normally but store?
                # For now, consume as statement tokens?
                # Or better: Extract locals here?
                # My ProcedureNode used to have .locals.
                # Now it needs to extract them from the body or I parse them here.
                # Let's parse them here and attach to a "DimNode" or just StatementNode?
                # The Analyzer will need to process Dim statements to add to scope.
                # I'll stick to StatementNode for Dim, but Analyzer must handle it.
                stmt = self.collect_statement()
                nodes.append(StatementNode(stmt))
                
            else:
                # Normal Statement
                stmt = self.collect_statement()
                if stmt:
                    nodes.append(StatementNode(stmt))
                else:
                    # Could be empty line
                    if self.current_token.type == 'NEWLINE':
                        self.advance()

        return nodes

    def collect_statement(self):
        tokens = []
        while self.current_token.type != 'NEWLINE' and self.current_token.type != 'EOF':
             # Check for "End" keywords that might terminate block unexpectedly (error case) or correctly?
             # No, parse_block check handles stopping condition.
             tokens.append(self.current_token)
             self.advance()
        if self.current_token.type == 'NEWLINE':
            self.advance()
        return tokens

    def parse_udt(self, module):
        self.consume('IDENTIFIER', 'Type')
        type_name = self.current_token.value
        self.advance()
        self.consume_statement()
        while self.current_token.type != 'EOF':
            if self.match('IDENTIFIER', 'End') and self.peek().value.lower() == 'type':
                self.advance()
                self.advance()
                self.consume_statement()
                break
            self.consume_statement()
