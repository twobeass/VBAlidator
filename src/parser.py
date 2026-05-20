from .lexer import Token

class Node:
    pass

class VariableNode(Node):
    def __init__(self, name, type_name, scope='Private', is_optional=False, is_paramarray=False, mechanism='ByRef', is_const=False, is_enum_member=False):
        self.name = name
        self.type_name = type_name
        self.scope = scope # Dim (Local), Private, Public, Global
        self.is_optional = is_optional
        self.is_paramarray = is_paramarray
        self.mechanism = mechanism
        self.is_const = is_const
        # `parse_enum` reuses VariableNode for each member; this flag lets
        # the analyzer register them with kind='EnumItem' (compile-time
        # constants, valid as `Const X = MyEnum.Member` RHS) rather than
        # generic Variable.
        self.is_enum_member = is_enum_member

    def __repr__(self):
        decl = "Const " if self.is_const else ""
        return f"{decl}Var({self.name} As {self.type_name} [{self.mechanism}])"

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
    def __init__(self, name, proc_type, return_type='Variant', scope='Public', is_declare=False, lib_name=None, alias_name=None, is_ptrsafe=False):
        self.name = name
        self.proc_type = proc_type # Sub, Function, Property Get/Set/Let
        self.return_type = return_type
        self.scope = scope
        self.is_declare = is_declare
        self.is_ptrsafe = is_ptrsafe
        self.lib_name = lib_name
        self.alias_name = alias_name
        self.args = [] # List of VariableNode
        self.locals = [] # List of VariableNode
        self.body = [] # List of nodes (StatementNode, WithNode)
        self.line = 0 # 1-based source line of the declaration (diagnostics)

    def __repr__(self):
        decl = "Declare " if self.is_declare else ""
        ptr = "PtrSafe " if self.is_ptrsafe else ""
        return f"{decl}{ptr}{self.proc_type} {self.name}() As {self.return_type}"

class TypeNode(Node):
    def __init__(self, name, scope='Public', is_enum=False):
        self.name = name
        self.scope = scope
        self.members = [] # List of VariableNode
        # `parse_enum` reuses TypeNode for simplicity; this flag lets the
        # analyzer distinguish "Enum of Long values" from a regular UDT —
        # callers passing an Enum-typed variable to a `ByRef p As Long`
        # are valid VBA (enums are Long under the hood).
        self.is_enum = is_enum

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
        # Phase 2.9 — DefInt/DefBool/DefStr… implicit typing per letter.
        # Lowercased single-letter ('a'..'z') → type name (Integer, Long, …).
        self.def_type_map = {}
        # Phase 3.6 — Option statements observed at module level.
        self.options = {
            "explicit": False,        # Option Explicit
            "compare": None,           # 'binary' | 'text' | 'database'
            "base": 0,                 # Option Base 0 | 1
            "private_module": False,   # Option Private Module
        }
        # Phase 3.1 — list of interface names this class/form implements.
        self.implements = []

class IfNode(Node):
    def __init__(self, condition_tokens, true_block, else_blocks=None, else_block=None):
        self.condition_tokens = condition_tokens
        self.true_block = true_block
        self.else_blocks = else_blocks if else_blocks else [] # List of (condition_tokens, block)
        self.else_block = else_block


class ForNode(Node):
    """For i = a To b [Step c]    OR    For Each x In coll"""
    def __init__(self, kind, var_token, header_tokens, body, line=0,
                 next_var_token=None, next_multi=False):
        self.kind = kind            # 'counter' | 'each'
        self.var_token = var_token  # Token for the loop variable name (may be None)
        self.header_tokens = header_tokens
        self.body = body
        self.line = line
        # Optional control variable on the closing `Next i` (None for bare
        # `Next`). `next_multi` flags the `Next j, i` multi-close form, where
        # the variable need not match this loop's counter.
        self.next_var_token = next_var_token
        self.next_multi = next_multi


class DoNode(Node):
    """Do [While|Until cond] ... Loop [While|Until cond]    AND    While ... Wend"""
    def __init__(self, condition_tokens, body, line=0, kind='do', condition_position='top'):
        self.condition_tokens = condition_tokens
        self.body = body
        self.line = line
        self.kind = kind  # 'do' | 'while_wend'
        # 'top' = condition checked before body (Do While/Until/While); 'bottom' = post-test (Loop While/Until); 'none' = unconditional Do/Loop
        self.condition_position = condition_position


class CaseClauseNode(Node):
    """One arm of a Select Case construct."""
    def __init__(self, header_tokens, body, is_else=False):
        self.header_tokens = header_tokens   # Token list after `Case` (without leading 'Case')
        self.body = body                      # List of nodes
        self.is_else = is_else


class SelectNode(Node):
    """Select Case <expr> ... End Select"""
    def __init__(self, expr_tokens, cases, line=0):
        self.expr_tokens = expr_tokens
        self.cases = cases                   # List[CaseClauseNode]
        self.line = line


class RedimNode(Node):
    """ReDim [Preserve] target(...) [As Type] [, ...]"""
    def __init__(self, preserve, targets, raw_tokens, line=0):
        self.preserve = preserve
        self.targets = targets               # List of (name_token, dim_tokens, as_type_or_None)
        self.raw_tokens = raw_tokens
        self.line = line


class EraseNode(Node):
    """Erase target1, target2, ..."""
    def __init__(self, targets, raw_tokens, line=0):
        self.targets = targets               # List of name tokens
        self.raw_tokens = raw_tokens
        self.line = line

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
        self.errors = []  # collected syntax errors (dicts)
        # `Next j, i` closes several nested For loops at once. The innermost
        # parse_for consumes the whole list and records how many *enclosing*
        # loops it also closed here, so those loops don't misfire VBA110.
        self._pending_next_closes = 0
        self.advance()

    def _record_syntax_error(self, message, line=None, rule_id="VBA_SYN001"):
        self.errors.append({
            "file": self.filename,
            "line": line if line is not None else self.current_token.line,
            "rule_id": rule_id,
            "severity": "error",
            "message": message,
        })

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
        # VBA182 — DefXxx statements must precede every variable / constant /
        # procedure / type declaration. Flips once the first such declaration
        # is parsed at module scope.
        self._module_decl_seen = False

        while self.current_token.type != 'EOF':
            if self.match('IDENTIFIER', 'Attribute'):
                self.parse_attribute(module)
            elif self.match('IDENTIFIER', 'VERSION'):
                # `.cls` / `.frm` file header — `VERSION 1.0 CLASS`
                # appears before the Attribute block and isn't a
                # language construct.
                self.consume_statement()
            elif self.match('IDENTIFIER', 'BEGIN'):
                # `.cls` BEGIN…END attribute block — consume every line
                # until a matching `END`. Cannot use consume_statement
                # because each property line ends with NEWLINE not `:`.
                self._consume_begin_end_block()
            elif self.match('IDENTIFIER', 'Option'):
                self._parse_option(module)
            elif self.match('IDENTIFIER', 'Implements'):
                self._parse_implements(module)
            elif self.current_token.type == 'IDENTIFIER' and self.current_token.value.lower() in (
                'defbool', 'defbyte', 'defint', 'deflng', 'deflnglng', 'deflngptr',
                'defcur', 'defsng', 'defdbl', 'defdate', 'defstr', 'defobj', 'defvar',
            ):
                self._parse_def_type(module)
            elif self.match('IDENTIFIER', 'Public') or self.match('IDENTIFIER', 'Private') or self.match('IDENTIFIER', 'Friend') or self.match('IDENTIFIER', 'Dim') or self.match('IDENTIFIER', 'Const') or self.match('IDENTIFIER', 'Global'):
                self._module_decl_seen = True
                self.parse_declaration(module)
            elif self.match('IDENTIFIER', 'Sub') or self.match('IDENTIFIER', 'Function') or self.match('IDENTIFIER', 'Property'):
                self._module_decl_seen = True
                self.procedures_parse(module, 'Public')
            elif self.match('IDENTIFIER', 'Type'):
                self._module_decl_seen = True
                self.parse_udt(module)
            elif self.match('IDENTIFIER', 'Event'):
                self._module_decl_seen = True
                # Handle implicit public Event
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
                self._module_decl_seen = True
                self.parse_enum(module, 'Public')
            elif self.match('NEWLINE'):
                self.advance()
            else:
                # P3.5 — anything reaching this branch is a token at
                # module level that didn't match a valid declaration form.
                # That means it's either an executable statement (illegal
                # outside a procedure body in VBA) or an unrecognised
                # token. Flag once per offending line, then skip so the
                # rest of the module continues to parse.
                self._record_module_level_executable(module)
                self.consume_statement()

        return module

    def _consume_begin_end_block(self):
        """Skip a `.cls` / `.frm` BEGIN…END attribute block.

        These appear in VBE-exported class/form files and look like:

            BEGIN
              MultiUse = -1  'True
              Persistable = 0
              DataBindingBehavior = 0
            END

        They are an MS-VBA serialisation artefact, not VBA code. The
        roundtrip path already strips them in `_strip_export_directives`;
        the parser now handles them here for the non-roundtrip path.
        """
        self.advance()  # consume BEGIN
        depth = 1
        while self.current_token.type != 'EOF' and depth > 0:
            if self.match('IDENTIFIER', 'BEGIN'):
                depth += 1
                self.advance()
            elif self.match('IDENTIFIER', 'END'):
                depth -= 1
                self.advance()
            else:
                self.advance()

    def _record_module_level_executable(self, module):
        """Emit VBA361 for executable code that appears at module level.

        Skips obvious non-statements (a stray standalone keyword the
        analyzer already complains about, a comment, …) so the rule
        only fires on real placement bugs.
        """
        tok = self.current_token
        if tok.type not in ('IDENTIFIER', 'OPERATOR'):
            return
        if tok.type == 'OPERATOR' and tok.value in (':',):
            return
        # Don't double-report on tokens that are part of a control flow
        # form parsed elsewhere (End / Loop / Wend / Next at module
        # level just unwind into the catch-all branch).
        if tok.type == 'IDENTIFIER' and tok.value.lower() in (
            'end', 'loop', 'wend', 'next', 'else', 'elseif', 'case',
        ):
            return
        self.errors.append({
            "file": self.filename,
            "line": tok.line,
            "rule_id": "VBA361",
            "severity": "error",
            "message": (
                f"Executable statement '{tok.value}' at module level. "
                f"Move it into a Sub or Function body, or wrap it in "
                f"`Public Const` / `Public Sub` / `Public Function`."
            ),
        })

    def _is_misplaced_statement_in_typelike(self):
        """Heuristic for VBA180/VBA181 — detect a statement wrongly placed
        inside a `Type`/`Enum` block.

        A valid member is `Name As Type`, `Name(dims) As …`, `Name = value`
        (Enum) or a bare `Name`. VBA is permissive about member names — many
        keywords (`Type`, `Name`, `Date`, …) are legal field names — so we do
        NOT match on a keyword blacklist. Instead we use the declaration
        *shape*: a line that starts with two consecutive identifiers where the
        second is not `As` (e.g. `Dim x …`, `Sub Foo`, `Static y`, `Set a`)
        cannot be a member declaration and is therefore a misplaced statement.
        """
        if self.current_token.type != 'IDENTIFIER':
            return False
        nxt = self.peek()
        return nxt.type == 'IDENTIFIER' and nxt.value.lower() != 'as'

    _DEFTYPE_TO_TYPE = {
        'defbool': 'Boolean', 'defbyte': 'Byte', 'defint': 'Integer',
        'deflng': 'Long', 'deflnglng': 'LongLong', 'deflngptr': 'LongPtr',
        'defcur': 'Currency', 'defsng': 'Single', 'defdbl': 'Double',
        'defdate': 'Date', 'defstr': 'String', 'defobj': 'Object',
        'defvar': 'Variant',
    }

    def _parse_def_type(self, module):
        """Parse `DefInt A-K, X` etc. and update module.def_type_map."""
        def_line = self.current_token.line
        keyword = self.current_token.value.lower()
        target_type = self._DEFTYPE_TO_TYPE.get(keyword, 'Variant')
        self.advance()  # consume DefXxx

        # VBA182 — DefXxx must precede all declarations.
        if getattr(self, '_module_decl_seen', False):
            self._record_syntax_error(
                "Compile error: Deftype statements must precede declarations",
                line=def_line, rule_id="VBA182")

        dup_reported = False
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            if self.current_token.type == 'IDENTIFIER' and len(self.current_token.value) >= 1:
                first = self.current_token.value[0].lower()
                last = first
                self.advance()
                if self.match('OPERATOR', '-'):
                    self.advance()
                    if self.current_token.type == 'IDENTIFIER' and len(self.current_token.value) >= 1:
                        last = self.current_token.value[0].lower()
                        self.advance()
                # Map every letter in the inclusive range
                if first.isalpha() and last.isalpha():
                    lo = min(ord(first), ord(last))
                    hi = max(ord(first), ord(last))
                    for code in range(lo, hi + 1):
                        letter = chr(code)
                        # VBA183 — a letter already covered by a previous
                        # DefXxx statement is a duplicate definition.
                        if letter in module.def_type_map and not dup_reported:
                            self._record_syntax_error(
                                "Compile error: Duplicate definition (Deftype letter range overlaps)",
                                line=def_line, rule_id="VBA183")
                            dup_reported = True
                        module.def_type_map[letter] = target_type
            elif self.match('OPERATOR', ','):
                self.advance()
            else:
                self.advance()
        self.consume_statement()

    def _parse_option(self, module):
        """Parse `Option Explicit | Compare {Binary|Text|Database} | Base N | Private Module`."""
        opt_line = self.current_token.line
        self.advance()  # consume 'Option'
        if self.current_token.type != 'IDENTIFIER':
            self.consume_statement()
            return
        kind = self.current_token.value.lower()
        self.advance()
        # VBA184 — the same Option statement must not appear twice.
        already = (
            (kind == 'explicit' and module.options.get('explicit'))
            or (kind == 'compare' and module.options.get('compare') is not None)
            or (kind == 'base' and getattr(module, '_option_base_seen', False))
            or (kind == 'private' and module.options.get('private_module'))
        )
        if already:
            self._record_syntax_error(
                f"Compile error: Duplicate Option statement (Option {kind.capitalize()})",
                line=opt_line, rule_id="VBA184")
        if kind == 'explicit':
            module.options['explicit'] = True
        elif kind == 'compare':
            if self.current_token.type == 'IDENTIFIER':
                module.options['compare'] = self.current_token.value.lower()
                self.advance()
        elif kind == 'base':
            module._option_base_seen = True
            if self.current_token.type == 'INTEGER':
                try:
                    module.options['base'] = int(self.current_token.value)
                except ValueError:
                    # Malformed `Option Base <not-an-int>` — leave the
                    # default (0). The lexer's INTEGER regex should make
                    # this unreachable in practice, but we keep the
                    # guard so a corrupted stream cannot crash the parser.
                    pass
                self.advance()
        elif kind == 'private':
            # `Option Private Module`
            if self.match('IDENTIFIER', 'Module'):
                self.advance()
                module.options['private_module'] = True
        self.consume_statement()

    def _parse_implements(self, module):
        """Parse `Implements <Interface[.SubInterface]>`."""
        self.advance()  # consume 'Implements'
        parts = []
        while self.current_token.type == 'IDENTIFIER':
            parts.append(self.current_token.value)
            self.advance()
            if self.match('OPERATOR', '.'):
                parts.append('.')
                self.advance()
            else:
                break
        if parts:
            module.implements.append("".join(parts))
        self.consume_statement()

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
            event_name_line = self.current_token.line
            if self.current_token.type == 'IDENTIFIER':
                event_name = self.current_token.value
                self.advance()

            proc = ProcedureNode(event_name, 'Event', scope=scope)
            proc.line = event_name_line

            if self.match('OPERATOR', '('):
                self.parse_arg_list(proc)

            self.consume_statement()
            module.procedures.append(proc)
            return

        # Handle Declare
        # [Public|Private] Declare [PtrSafe] Sub/Function ...
        if self.match('IDENTIFIER', 'Declare'):
            declare_line = self.current_token.line
            self.advance() # consume Declare

            # Optional PtrSafe — required on 64-bit Office hosts.
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

            proc = ProcedureNode(
                proc_name, proc_type,
                scope=scope, is_declare=True,
                lib_name=lib_name, alias_name=alias_name,
                is_ptrsafe=is_ptrsafe,
            )
            proc.declare_line = declare_line  # used for diagnostics
            proc.line = declare_line

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
        is_const = False
        if scope.lower() == 'const':
            # Bare module-level `Const X = …` (no Public/Private prefix). The
            # `Const` keyword was consumed as `scope`; default the real scope
            # to Private (VBA's default for an unqualified module constant).
            is_const = True
            scope = 'Private'
        elif scope.lower() in ('public', 'private', 'global', 'friend'):
             if self.match('IDENTIFIER', 'Const'):
                 is_const = True
                 self.advance()

        # Check if WithEvents
        if self.match('IDENTIFIER', 'WithEvents'):
            self.advance()

        # Dim x [(dims)] As Type [= init]
        while True:
            if self.current_token.type == 'IDENTIFIER':
                var_name = self.current_token.value
                self.advance()

                # P2.6 — Array suffix `name(...)` precedes `As` in VBA.
                # Parsing it first preserves the declared element type so
                # `cells(0).val` resolves to the real member type instead
                # of degrading to Variant.
                is_array = False
                if self.match('OPERATOR', '('):
                    is_array = True
                    while self.current_token.type != 'EOF' and not self.match('OPERATOR', ')'):
                        self.advance()
                    self.consume('OPERATOR', ')')

                var_type = 'Variant'
                if self.match('IDENTIFIER', 'As'):
                    self.advance()
                    var_type = self.parse_type_signature()

                if is_array:
                    var_type += "()"

                if self.match('OPERATOR', '='):
                     while self.current_token.type not in ('NEWLINE', 'EOF') and not self.match('OPERATOR', ','):
                         self.advance()

                module.variables.append(VariableNode(var_name, var_type, scope, is_const=is_const))
            
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
        proc_line = self.current_token.line
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
        proc.line = proc_line
        
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

        # Ensure we consumed End Sub/Function/Property. AI generators
        # occasionally close a Function with `End Sub` (or vice versa);
        # VBE rejects this at compile time. Surface it as VBA350 so the
        # mismatch is caught before VBE ever sees it.
        consumed_end = False
        if self.match('IDENTIFIER', 'End'):
             end_line = self.current_token.line
             self.advance()
             actual = self.current_token.value.lower() if self.current_token.type == 'IDENTIFIER' else None
             if actual == end_marker:
                 self.advance()
                 consumed_end = True
             elif actual in ('sub', 'function', 'property'):
                 # Wrong terminator keyword (`End Sub` closing a Function
                 # etc.). Consume it so recovery continues cleanly.
                 self.errors.append({
                     "file": self.filename,
                     "line": end_line,
                     "rule_id": "VBA350",
                     "severity": "error",
                     "message": (
                         f"`End {self.current_token.value}` closes "
                         f"{proc_type} '{proc_name}'. Use "
                         f"`End {end_marker.capitalize()}` to match the "
                         f"procedure kind."
                     ),
                 })
                 self.advance()

        # VBA151 — only comments may follow `End Sub/Function/Property`.
        if consumed_end:
            while self.current_token.type not in ('NEWLINE', 'EOF'):
                if self.current_token.type == 'COMMENT':
                    break
                if self.current_token.type == 'OPERATOR' and self.current_token.value == ':':
                    self.advance()
                    continue
                self._record_syntax_error(
                    "Compile error: Only comments may appear after End Sub, "
                    "End Function, or End Property",
                    line=self.current_token.line, rule_id="VBA151")
                break
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
                    parts = marker.lower().split()
                    if parts[0] != val:
                        continue
                    # Multi-word markers ("End If", "End With", ...) are handled
                    # by the dedicated `End` branch above which requires the full
                    # two-token form. Don't prefix-match them here — that would
                    # consume the standalone `End` *statement* (the program
                    # terminator) and trap callers in an infinite loop when the
                    # If/With body contains one.
                    if len(parts) > 1:
                        continue
                    # Potential match.
                    # If marker is "Next", and we have "Next i", it's a match.
                    # If marker is "Else", and we have "Else", it's a match.
                    return nodes

                # VALIDATION: Check for unexpected block terminators. These
                # are the *inverse* of the missing-terminator checks in the
                # individual block parsers: here a terminator appears with no
                # matching opener in scope.
                _INVERSE_TERMINATORS = {
                    'next': ("Compile error: Next without For", "VBA116"),
                    'loop': ("Compile error: Loop without Do", "VBA117"),
                    'wend': ("Compile error: Wend without While", "VBA118"),
                    'else': ("Compile error: Else without If", "VBA119"),
                    'elseif': ("Compile error: ElseIf without If", "VBA119"),
                }
                if val in _INVERSE_TERMINATORS:
                    msg, rid = _INVERSE_TERMINATORS[val]
                    self._record_syntax_error(msg, rule_id=rid)
                    # We consume it to avoid infinite loop, but it's an error
                    self.consume_statement()
                    continue

                if val == 'end':
                    peek_val = self.peek().value.lower()
                    if peek_val in ('if', 'select', 'with', 'function', 'sub', 'property'):
                        # Found End X that was NOT in end_markers -> Unexpected
                        self._record_syntax_error(
                            f"Compile error: End {self.peek().value} without matching block statement",
                            rule_id="VBA119",
                        )
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
                # A nested `Next j, i` closed this loop too; stop the body here
                # and let the enclosing parse_for account for the close.
                if self._pending_next_closes > 0:
                    return nodes

            elif self.match('IDENTIFIER', 'Do'):
                nodes.append(self.parse_do())
                
            elif self.match('IDENTIFIER', 'Select'):
                nodes.append(self.parse_select())

            elif self.match('IDENTIFIER', 'While'):
                 nodes.append(self.parse_while())

            elif self.match('IDENTIFIER', 'Dim') or self.match('IDENTIFIER', 'Static'):
                stmt = self.collect_statement()
                nodes.append(StatementNode(stmt))

            elif self._matches_module_only_keyword():
                # P3.5 — `Type`, `Enum`, `Declare`, `Option`, `Implements`
                # and the `DefXxx` family are module-level-only in VBA.
                # Flag once and consume the line so the procedure body
                # keeps parsing.
                self._record_proc_level_module_only(nodes)

            elif self.match('IDENTIFIER', 'ReDim'):
                nodes.append(self.parse_redim())

            elif self.match('IDENTIFIER', 'Erase'):
                nodes.append(self.parse_erase())

            elif self.match('IDENTIFIER', 'Attribute'):
                # Ignore attribute statements inside blocks
                self.consume('IDENTIFIER', 'Attribute')
                self.consume_statement()
                
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
        line = self.current_token.line
        self.consume('IDENTIFIER', 'While')
        condition_tokens = self.collect_statement()  # Everything until newline

        body = self.parse_block(end_markers=["Wend"])

        if not self.consume('IDENTIFIER', 'Wend'):
            self._record_syntax_error(
                "Compile error: While without Wend", line=line, rule_id="VBA112")
        self.consume_statement()

        return DoNode(
            condition_tokens=condition_tokens,
            body=body,
            line=line,
            kind='while_wend',
            condition_position='top',
        )

    def parse_with(self):
        line = self.current_token.line
        self.consume('IDENTIFIER', 'With')
        expr_tokens = []
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            expr_tokens.append(self.current_token)
            self.advance()
        self.consume_statement()

        body = self.parse_block(end_markers=["End With"])

        if not self.consume('IDENTIFIER', 'End') or not self.consume('IDENTIFIER', 'With'):
            self._record_syntax_error(
                "Compile error: Expected End With", line=line, rule_id="VBA113")
        else:
            self.consume_statement()

        return WithNode(expr_tokens, body)

    def parse_if_stmt(self):
        # If <condition> Then <newline> [Block]
        # If <condition> Then <statement> [Else <statement>] [newline] [Single Line]

        if_line = self.current_token.line
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
             self._record_syntax_error(
                 "Syntax Error: Missing 'Then' after If condition."
             )
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
             end_if_seen = False

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
                             end_if_seen = True
                         break

                     else:
                         break
                 else:
                     break

             if not end_if_seen:
                 self._record_syntax_error(
                     "Compile error: Block If without End If",
                     line=if_line, rule_id="VBA115")

             return IfNode(condition_tokens, true_block, else_blocks, else_block)
             
        else:
             # Single Line If: If <cond> Then <stmt[: stmt]>* [Else <stmt[: stmt]>*]
             # Bug fix: previously this branch fell through and returned None,
             # which caused condition + body tokens to be silently discarded
             # (no analyzer pass on them). Now we build a proper IfNode so the
             # analyzer walks the condition and each body statement.
             true_block = self._collect_inline_block(stop_on_else=True)
             else_block = None
             if self.match('IDENTIFIER', 'Else'):
                 self.advance()
                 else_block = self._collect_inline_block(stop_on_else=False)

             if self.match('NEWLINE'):
                 self.advance()

             return IfNode(condition_tokens, true_block, [], else_block)

    def _collect_inline_block(self, stop_on_else):
        """Collect colon-separated statements on the current line.

        Used for single-line If bodies. Unlike :meth:`collect_statement` this
        also breaks on the ``Else`` keyword so that ``If ... Then a Else b``
        parses cleanly into separate true / else blocks.
        """
        block = []
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            if stop_on_else and self.match('IDENTIFIER', 'Else'):
                break
            stmt_tokens = []
            while self.current_token.type not in ('NEWLINE', 'EOF'):
                if stop_on_else and self.match('IDENTIFIER', 'Else'):
                    break
                if self.current_token.type == 'OPERATOR' and self.current_token.value == ':':
                    stmt_tokens.append(self.current_token)
                    self.advance()
                    break
                stmt_tokens.append(self.current_token)
                self.advance()
            if stmt_tokens:
                block.append(StatementNode(stmt_tokens))
        return block



    def parse_for(self):
        line = self.current_token.line
        self.consume('IDENTIFIER', 'For')

        kind = 'counter'
        var_token = None
        header_tokens = []

        if self.match('IDENTIFIER', 'Each'):
            kind = 'each'
            self.advance()  # consume 'Each'
            if self.current_token.type == 'IDENTIFIER':
                var_token = self.current_token

        else:
            if self.current_token.type == 'IDENTIFIER':
                var_token = self.current_token

        # Capture the header tokens up to NEWLINE so the analyzer can walk
        # the iteration expression (range, collection, step, …).
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            header_tokens.append(self.current_token)
            self.advance()
        self.consume_statement()

        body = self.parse_block(end_markers=["Next"])

        next_var_token = None
        next_multi = False
        if not self.consume('IDENTIFIER', 'Next'):
            # An enclosing loop's `Next j, i` may already have closed this one.
            if self._pending_next_closes > 0:
                self._pending_next_closes -= 1
            else:
                self._record_syntax_error(
                    "Compile error: For without Next", line=line, rule_id="VBA110")
        elif self.current_token.type == 'IDENTIFIER':
            # Optional variable name after Next (`Next i`).
            next_var_token = self.current_token
            self.advance()
            # `Next j, i [, ...]` closes several nested loops on one line. Each
            # extra name closes one enclosing loop; record that so those
            # parse_for frames don't report a missing Next.
            extra_closes = 0
            while self.match('OPERATOR', ','):
                next_multi = True
                self.advance()  # consume ','
                if self.current_token.type == 'IDENTIFIER':
                    extra_closes += 1
                    self.advance()
            if extra_closes:
                self._pending_next_closes += extra_closes
        self.consume_statement()

        return ForNode(
            kind=kind,
            var_token=var_token,
            header_tokens=header_tokens,
            body=body,
            line=line,
            next_var_token=next_var_token,
            next_multi=next_multi,
        )

    def parse_do(self):
        line = self.current_token.line
        self.consume('IDENTIFIER', 'Do')

        # Optional top-tested condition: Do While <cond>  /  Do Until <cond>
        condition_tokens = []
        condition_position = 'none'
        if self.match('IDENTIFIER', 'While') or self.match('IDENTIFIER', 'Until'):
            self.advance()  # consume While/Until
            condition_position = 'top'
            while self.current_token.type not in ('NEWLINE', 'EOF'):
                condition_tokens.append(self.current_token)
                self.advance()
        else:
            # Skip rest of header line (could be just `Do` followed by comment)
            while self.current_token.type not in ('NEWLINE', 'EOF'):
                self.advance()
        self.consume_statement()

        body = self.parse_block(end_markers=["Loop"])

        if not self.consume('IDENTIFIER', 'Loop'):
            self._record_syntax_error(
                "Compile error: Do without Loop", line=line, rule_id="VBA111")
        # Optional bottom-tested condition: Loop While <cond>  /  Loop Until <cond>
        if self.match('IDENTIFIER', 'While') or self.match('IDENTIFIER', 'Until'):
            self.advance()
            if condition_position == 'none':
                condition_position = 'bottom'
            while self.current_token.type not in ('NEWLINE', 'EOF'):
                condition_tokens.append(self.current_token)
                self.advance()
        else:
            while self.current_token.type not in ('NEWLINE', 'EOF'):
                self.advance()
        self.consume_statement()

        return DoNode(
            condition_tokens=condition_tokens,
            body=body,
            line=line,
            kind='do',
            condition_position=condition_position,
        )

    def parse_select(self):
        line = self.current_token.line
        self.consume('IDENTIFIER', 'Select')
        self.consume('IDENTIFIER', 'Case')

        # Capture the selector expression up to NEWLINE.
        expr_tokens = []
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            expr_tokens.append(self.current_token)
            self.advance()
        self.consume_statement()

        # Skip stray newlines / comments before first Case.
        cases = []
        while self.current_token.type != 'EOF':
            if self.match('NEWLINE') or self.match('COMMENT'):
                self.advance()
                continue

            # End of select?
            if self.match('IDENTIFIER', 'End'):
                peek_val = self.peek().value.lower() if self.peek() else ''
                if peek_val == 'select':
                    break

            if self.match('IDENTIFIER', 'Case'):
                self.advance()
                is_else = False
                header_tokens = []

                if self.match('IDENTIFIER', 'Else'):
                    is_else = True
                    self.advance()
                else:
                    while self.current_token.type not in ('NEWLINE', 'EOF'):
                        header_tokens.append(self.current_token)
                        self.advance()
                self.consume_statement()

                # Body of this case ends at the next Case / End Select.
                case_body = self.parse_block(end_markers=["Case", "End Select"])
                cases.append(CaseClauseNode(header_tokens, case_body, is_else=is_else))
            else:
                # Defensive: avoid infinite loop on malformed select.
                self._record_syntax_error(
                    f"Syntax Error: Expected 'Case' inside Select block, got '{self.current_token.value}'."
                )
                self.consume_statement()

        if not self.consume('IDENTIFIER', 'End') or not self.consume('IDENTIFIER', 'Select'):
            self._record_syntax_error(
                "Compile error: Select Case without End Select",
                line=line, rule_id="VBA114")
        else:
            self.consume_statement()

        return SelectNode(expr_tokens=expr_tokens, cases=cases, line=line)

    # ---- P3.5 — module-only keyword detection at procedure level ----
    # Keywords that VBA only accepts at module level. We add the `DefXxx`
    # family at runtime via prefix match.
    _MODULE_ONLY_KEYWORDS = {
        'type', 'enum', 'declare', 'option', 'implements',
    }

    def _matches_module_only_keyword(self) -> bool:
        if self.current_token.type != 'IDENTIFIER':
            return False
        low = self.current_token.value.lower()
        if low in self._MODULE_ONLY_KEYWORDS:
            return True
        # `DefBool`, `DefInt`, `DefStr`, … — all start with `def` and
        # are 6+ chars; differentiate from local variable names like
        # `defaultValue` by requiring the 4th character to be uppercase
        # in the original source (DefInt vs default*) — too fragile,
        # we use a hardcoded set instead.
        return low in {
            'defbool', 'defbyte', 'defint', 'deflng', 'deflnglng',
            'deflngptr', 'defcur', 'defsng', 'defdbl', 'defdate',
            'defstr', 'defobj', 'defvar',
        }

    def _record_proc_level_module_only(self, nodes):
        tok = self.current_token
        self.errors.append({
            "file": self.filename,
            "line": tok.line,
            "rule_id": "VBA360",
            "severity": "error",
            "message": (
                f"`{tok.value}` is only legal at module level — VBA does "
                f"not allow it inside a procedure body. Move the "
                f"declaration to module scope (outside any "
                f"Sub/Function/Property)."
            ),
        })
        # Wrap the offender as a StatementNode so analyse / reporting
        # paths don't trip on a missing node, then advance past the line.
        stmt = self.collect_statement()
        if stmt:
            nodes.append(StatementNode(stmt))

    def parse_redim(self):
        """ReDim [Preserve] target1(dims) [As Type] [, target2(...) ...]"""
        line = self.current_token.line
        raw_tokens = []
        self.consume('IDENTIFIER', 'ReDim')

        preserve = False
        if self.match('IDENTIFIER', 'Preserve'):
            preserve = True
            raw_tokens.append(self.current_token)
            self.advance()

        targets = []
        # Collect token-stream for simple ReDim parsing.
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            if self.current_token.type == 'OPERATOR' and self.current_token.value == ':':
                raw_tokens.append(self.current_token)
                self.advance()
                break

            # Leading-dot target — `ReDim .points(1 To 10)` inside a
            # `With X` block. We preserve the leading `.` in chain_tokens
            # so the analyzer can substitute the with-stack anchor.
            leading_dot = None
            if self.match('OPERATOR', '.'):
                leading_dot = self.current_token
                raw_tokens.append(leading_dot)
                self.advance()

            # Target name token
            if self.current_token.type != 'IDENTIFIER':
                raw_tokens.append(self.current_token)
                self.advance()
                continue

            name_token = self.current_token
            raw_tokens.append(name_token)
            self.advance()

            # Track the full chain so the analyser can resolve dotted
            # ReDim targets (`ReDim This.scopes(1 To N)`) via the same
            # member-walker that powers P2.6 member-chain typing.
            chain_tokens = [leading_dot, name_token] if leading_dot else [name_token]
            while self.match('OPERATOR', '.'):
                dot_tok = self.current_token
                raw_tokens.append(dot_tok)
                chain_tokens.append(dot_tok)
                self.advance()
                if self.current_token.type == 'IDENTIFIER':
                    name_token = self.current_token
                    raw_tokens.append(name_token)
                    chain_tokens.append(name_token)
                    self.advance()

            # Dimension expression in parens
            dim_tokens = []
            if self.match('OPERATOR', '('):
                paren_depth = 0
                while self.current_token.type not in ('NEWLINE', 'EOF'):
                    raw_tokens.append(self.current_token)
                    if self.current_token.type == 'OPERATOR' and self.current_token.value == '(':
                        paren_depth += 1
                        if paren_depth > 1:
                            dim_tokens.append(self.current_token)
                    elif self.current_token.type == 'OPERATOR' and self.current_token.value == ')':
                        paren_depth -= 1
                        if paren_depth == 0:
                            self.advance()
                            break
                        dim_tokens.append(self.current_token)
                    else:
                        dim_tokens.append(self.current_token)
                    self.advance()

            # Optional 'As Type'
            as_type = None
            if self.match('IDENTIFIER', 'As'):
                raw_tokens.append(self.current_token)
                self.advance()
                # Capture type signature (best-effort, until comma/newline)
                type_tokens = []
                while self.current_token.type not in ('NEWLINE', 'EOF') \
                        and not (self.current_token.type == 'OPERATOR' and self.current_token.value in (',', ':')):
                    type_tokens.append(self.current_token)
                    raw_tokens.append(self.current_token)
                    self.advance()
                as_type = ''.join(t.value for t in type_tokens).strip() or None

            targets.append((name_token, dim_tokens, as_type, chain_tokens))

            # Comma → next target on same statement
            if self.match('OPERATOR', ','):
                raw_tokens.append(self.current_token)
                self.advance()
                continue
            break

        if self.current_token.type == 'NEWLINE':
            self.advance()

        return RedimNode(preserve=preserve, targets=targets, raw_tokens=raw_tokens, line=line)

    def parse_erase(self):
        """Erase target1, target2, ..."""
        line = self.current_token.line
        raw_tokens = []
        self.consume('IDENTIFIER', 'Erase')

        targets = []  # Either Token (bare name) or list[Token] (dotted chain).
        while self.current_token.type not in ('NEWLINE', 'EOF'):
            if self.current_token.type == 'OPERATOR' and self.current_token.value == ':':
                raw_tokens.append(self.current_token)
                self.advance()
                break

            if self.current_token.type == 'IDENTIFIER':
                first_tok = self.current_token
                raw_tokens.append(first_tok)
                self.advance()
                # Capture qualified name parts (foo.bar.baz). We store the
                # whole chain so the analyzer can walk it via
                # `_resolve_lhs_type` — bare-name targets become a list of
                # length 1, which keeps the existing single-name code path
                # working unchanged.
                chain = [first_tok]
                while self.match('OPERATOR', '.'):
                    dot_tok = self.current_token
                    raw_tokens.append(dot_tok)
                    chain.append(dot_tok)
                    self.advance()
                    if self.current_token.type == 'IDENTIFIER':
                        member_tok = self.current_token
                        raw_tokens.append(member_tok)
                        chain.append(member_tok)
                        self.advance()
                targets.append(chain)
            else:
                raw_tokens.append(self.current_token)
                self.advance()

            if self.match('OPERATOR', ','):
                raw_tokens.append(self.current_token)
                self.advance()
                continue

        if self.current_token.type == 'NEWLINE':
            self.advance()

        return EraseNode(targets=targets, raw_tokens=raw_tokens, line=line)

    def collect_statement(self, consume_newline=True):
        tokens = []
        while self.current_token.type != 'NEWLINE' and self.current_token.type != 'EOF':
             tokens.append(self.current_token)
             
             # Check for statement separator ':'
             if self.current_token.type == 'OPERATOR' and self.current_token.value == ':':
                 tokens.append(self.current_token) # Include colon for Label detection
                 self.advance()
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

            # VBA181 — a statement wrongly placed inside a Type block.
            if self._is_misplaced_statement_in_typelike():
                self._record_syntax_error(
                    "Compile error: Statement invalid inside Type block",
                    line=self.current_token.line, rule_id="VBA181")
                self.consume_statement()
                continue

            # Parse Member: Name [(dims)] As Type [* length]
            if self.current_token.type == 'IDENTIFIER':
                var_name = self.current_token.value
                self.advance()

                # P2.6 — Array suffix `name(...)` comes BEFORE `As` in VBA,
                # not after. Capturing it here lets the element type carry
                # forward into deep member-chains like `r.cells(0).val`.
                is_array = False
                if self.match('OPERATOR', '('):
                    is_array = True
                    while not self.match('OPERATOR', ')') and self.current_token.type != 'EOF':
                        self.advance()
                    self.consume('OPERATOR', ')')

                var_type = 'Variant'
                if self.match('IDENTIFIER', 'As'):
                    self.advance()
                    var_type = self.parse_type_signature()

                if is_array:
                    var_type += "()"

                if self.match('OPERATOR', '*'):
                    self.advance()
                    self.advance()

                udt.members.append(VariableNode(var_name, var_type, 'Public'))
            
            if self.match('OPERATOR', ':'):
                self.advance()
            else:
                self.consume_statement()
            
        module.types[type_name] = udt

    def parse_enum(self, module, scope='Public'):
        self.consume('IDENTIFIER', 'Enum')
        enum_name = self.current_token.value
        self.advance()
        self.consume_statement()

        # Enums are basically Longs with named constants
        # We need to register the Enum Type AND the Enum Members as global/module constants
        udt = TypeNode(enum_name, scope, is_enum=True) # Reuse TypeNode for simplicity

        while self.current_token.type != 'EOF':
            if self.match('IDENTIFIER', 'End') and self.peek().value.lower() == 'enum':
                self.advance()
                self.advance()
                self.consume_statement()
                break

            # VBA180 — a statement wrongly placed inside an Enum block.
            if self._is_misplaced_statement_in_typelike():
                self._record_syntax_error(
                    "Compile error: Statement invalid inside Enum block",
                    line=self.current_token.line, rule_id="VBA180")
                self.consume_statement()
                continue

            # Member: Name = Value
            if self.current_token.type == 'IDENTIFIER':
                member_name = self.current_token.value
                self.advance()

                # Enum members are constants.
                # We register the Enum Type in module.types
                # AND we register the members as module-level variables (Consts)

                var = VariableNode(member_name, 'Long', scope, is_enum_member=True) # Enum members are Long
                module.variables.append(var)
                udt.members.append(var)

                if self.match('OPERATOR', '='):
                    self.advance()
                    # Skip value
                    while self.current_token.type not in ('NEWLINE', 'EOF', 'COMMENT'):
                        if self.match('OPERATOR', ':'):
                            break
                        self.advance()

            if self.match('OPERATOR', ':'):
                self.advance()
            else:
                self.consume_statement()

        module.types[enum_name] = udt
