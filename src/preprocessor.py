"""VBA conditional-compilation preprocessor.

Filters the token stream based on `#If` / `#ElseIf` / `#Else` /
`#End If` and `#Const` directives. Symbol lookup is case-insensitive
(matching VBA's behaviour) and undefined identifiers evaluate to False
/ Empty.

Expression evaluation is sandboxed: we parse the directive expression
through Python's `ast` module and walk only a strict whitelist of
nodes (BoolOp, UnaryOp, Compare, Name, Constant). `eval()` is *not*
used — bandit B307 would rightly flag it as a code-execution risk
even with a stripped globals dict.
"""
import ast
import operator


class SafeDict(dict):
    """Case-insensitive defines lookup with `False` for misses."""

    def __init__(self, source):
        super().__init__()
        for k, v in source.items():
            super().__setitem__(k.upper(), v)

    def get_ci(self, key):
        upper = key.upper()
        if super().__contains__(upper):
            return super().__getitem__(upper)
        return False


# Whitelisted boolean and comparison operators.
_BOOL_OPS = {
    ast.And: lambda values: all(values),
    ast.Or: lambda values: any(values),
}

_UNARY_OPS = {
    ast.Not: operator.not_,
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
}

_CMP_OPS = {
    ast.Eq: operator.eq,
    ast.NotEq: operator.ne,
    ast.Lt: operator.lt,
    ast.LtE: operator.le,
    ast.Gt: operator.gt,
    ast.GtE: operator.ge,
}


class _UnsupportedExpression(Exception):
    """Raised when the directive expression contains a non-whitelisted node."""


def _safe_eval(node, env: SafeDict):
    """Recursively evaluate a parsed AST under a strict whitelist.

    Anything outside the whitelist (function calls, attribute access,
    subscripts, comprehensions, lambdas, …) raises and the caller
    treats the directive as False — matching VBA's behaviour for
    unparseable directives.
    """
    if isinstance(node, ast.Expression):
        return _safe_eval(node.body, env)

    if isinstance(node, ast.BoolOp):
        op_fn = _BOOL_OPS.get(type(node.op))
        if op_fn is None:
            raise _UnsupportedExpression(f"BoolOp {type(node.op).__name__}")
        return op_fn(_safe_eval(v, env) for v in node.values)

    if isinstance(node, ast.UnaryOp):
        op_fn = _UNARY_OPS.get(type(node.op))
        if op_fn is None:
            raise _UnsupportedExpression(f"UnaryOp {type(node.op).__name__}")
        return op_fn(_safe_eval(node.operand, env))

    if isinstance(node, ast.Compare):
        left = _safe_eval(node.left, env)
        for op, comparator in zip(node.ops, node.comparators):
            op_fn = _CMP_OPS.get(type(op))
            if op_fn is None:
                raise _UnsupportedExpression(f"Cmp {type(op).__name__}")
            right = _safe_eval(comparator, env)
            if not op_fn(left, right):
                return False
            left = right
        return True

    if isinstance(node, ast.Name):
        # Undefined → False (VBA semantic).
        return env.get_ci(node.id)

    if isinstance(node, ast.Constant):
        # Numbers, strings, booleans, None — all literal, no execution.
        return node.value

    raise _UnsupportedExpression(type(node).__name__)


class Preprocessor:
    def __init__(self, tokens, defines):
        self.tokens = tokens
        self.defines = defines
        self.stack = [{"active": True, "taken": False}] # Root scope

    def evaluate(self, tokens):
        # Convert tokens to a Python-parseable expression. VBA `=` becomes
        # `==`, `<>` becomes `!=`, And/Or/Not stay (they are already
        # Python keywords once lower-cased).
        expr = []
        for t in tokens:
            val = t.value
            low = val.lower()
            if low == 'and':
                val = ' and '
            elif low == 'or':
                val = ' or '
            elif low == 'not':
                val = ' not '
            elif val == '=':
                val = ' == '
            elif val == '<>':
                val = ' != '
            expr.append(val)

        expr_str = "".join(expr).strip()
        if not expr_str:
            return False

        env = SafeDict(self.defines)

        try:
            parsed = ast.parse(expr_str, mode="eval")
            return bool(_safe_eval(parsed.body, env))
        except (_UnsupportedExpression, SyntaxError, ValueError, TypeError):
            return False

    def process(self):
        iterator = iter(self.tokens)
        current_token = next(iterator, None)

        while current_token:
            if current_token.type == 'PREPROCESSOR':
                directive = current_token.value.lower()

                # Handle #If
                if directive == '#if':
                    # Collect condition until 'Then' or Newline
                    cond_tokens = []
                    current_token = next(iterator, None)
                    while current_token and current_token.type not in ('NEWLINE', 'EOF'):
                        if current_token.value.lower() == 'then':
                            current_token = next(iterator, None) # Skip Then
                            break
                        cond_tokens.append(current_token)
                        current_token = next(iterator, None)

                    # Evaluate
                    parent = self.stack[-1]
                    if parent["active"]:
                        result = self.evaluate(cond_tokens)
                        self.stack.append({"active": result, "taken": result})
                    else:
                        self.stack.append({"active": False, "taken": True}) # Parent inactive, so this is inactive

                # Handle #ElseIf
                elif directive == '#elseif':
                    cond_tokens = []
                    current_token = next(iterator, None)
                    while current_token and current_token.type not in ('NEWLINE', 'EOF'):
                         if current_token.value.lower() == 'then':
                            current_token = next(iterator, None)
                            break
                         cond_tokens.append(current_token)
                         current_token = next(iterator, None)

                    current_scope = self.stack[-1]
                    parent = self.stack[-2] # Parent of current #If

                    if parent["active"] and not current_scope["taken"]:
                        result = self.evaluate(cond_tokens)
                        current_scope["active"] = result
                        if result: current_scope["taken"] = True
                    else:
                        current_scope["active"] = False

                # Handle #Else
                elif directive == '#else':
                    current_scope = self.stack[-1]
                    parent = self.stack[-2]

                    if parent["active"] and not current_scope["taken"]:
                        current_scope["active"] = True
                        current_scope["taken"] = True
                    else:
                        current_scope["active"] = False

                    current_token = next(iterator, None) # Consume newline if present?

                # Handle #End If
                elif directive == '#end':
                    # Check next token for 'if'
                    next_tok = next(iterator, None)
                    if next_tok and next_tok.value.lower() == 'if':
                         self.stack.pop()
                         current_token = next(iterator, None)
                    else:
                         # Just #End? Unlikely in preprocessor, usually #End If.
                         # Treat as pop anyway? Or error?
                         # Assume it's #End If
                         self.stack.pop()
                         current_token = next_tok

                # Handle #Const
                elif directive == '#const':
                    # #Const Identifier = Expression
                    # We need to parse identifier
                    current_token = next(iterator, None)
                    if current_token and current_token.type == 'IDENTIFIER':
                         const_name = current_token.value
                         current_token = next(iterator, None)
                         if current_token and current_token.value == '=':
                              current_token = next(iterator, None)
                              # Parse expression until Newline
                              expr_tokens = []
                              while current_token and current_token.type not in ('NEWLINE', 'EOF'):
                                   expr_tokens.append(current_token)
                                   current_token = next(iterator, None)

                              # Evaluate and assign ONLY if active
                              if self.stack[-1]["active"]:
                                   val = self.evaluate(expr_tokens)
                                   self.defines[const_name.upper()] = val
                         else:
                              # Syntax error in #Const, skip line
                              while current_token and current_token.type not in ('NEWLINE', 'EOF'):
                                   current_token = next(iterator, None)
                    else:
                         # Syntax error
                         while current_token and current_token.type not in ('NEWLINE', 'EOF'):
                              current_token = next(iterator, None)

                else:
                    # Unknown directive, ignore or yield?
                    yield current_token
                    current_token = next(iterator, None)

                # Directives themselves are consumed.
                # If we are at a newline, yield it to keep line count?
                if current_token and current_token.type == 'NEWLINE':
                    yield current_token
                    current_token = next(iterator, None)

            else:
                # Normal token
                if self.stack[-1]["active"]:
                    yield current_token
                else:
                    # If inactive, we still yield newlines to preserve line numbers
                    if current_token.type == 'NEWLINE':
                        yield current_token

                current_token = next(iterator, None)
