class Preprocessor:
    def __init__(self, tokens, defines):
        self.tokens = tokens
        self.defines = defines
        self.stack = [{"active": True, "taken": False}] # Root scope

    def evaluate(self, tokens):
        # Convert tokens to string for evaluation
        # Replace VBA operators with Python equivalents
        expr = []
        for t in tokens:
            val = t.value
            if val.lower() == 'and': val = ' and '
            elif val.lower() == 'or': val = ' or '
            elif val.lower() == 'not': val = ' not '
            elif val == '=': val = '=='
            elif val == '<>': val = '!='
            expr.append(val)
        
        expr_str = "".join(expr)
        
        # Safe eval context
        context = self.defines.copy()
        # Ensure True/False exist
        context['True'] = True
        context['False'] = False
        
        try:
            # We must be careful with identifiers that are NOT in defines.
            # In VBA #If, undefined constants are usually Empty/0/False.
            # We can use a defaultdict or custom dict.
            class SafeDict(dict):
                def __missing__(self, key):
                    return False # Assume undefined is False
            
            return bool(eval(expr_str, {}, SafeDict(context)))
        except Exception as e:
            # print(f"Preprocessor Eval Error: {e} in {expr_str}")
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
