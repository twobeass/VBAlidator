import re

class Token:
    def __init__(self, type, value, line, column):
        self.type = type
        self.value = value
        self.line = line
        self.column = column

    def __repr__(self):
        return f"Token({self.type}, {repr(self.value)}, Line:{self.line})"

class Lexer:
    def __init__(self, code):
        self.code = code
        self.pos = 0
        self.line = 1
        self.column = 1

        # Regex patterns
        self.token_specs = [
            ('COMMENT', r"'.*"),
            ('STRING', r'"(""|[^"])*"'),
            ('PREPROCESSOR', r'#[a-zA-Z_]\w*'),
            ('DATELITERAL', r'\#[^#\r\n]+\#'),
            ('HEX', r'&H[0-9A-Fa-f]+'),
            ('FLOAT', r'\d+\.\d+'),
            ('INTEGER', r'\d+'),
            ('LINE_CONTINUATION', r'[ \t]+_(\r\n|\n)'), # Handle line continuation
            ('NEWLINE', r'(\r\n|\n)'), # Removed : from newline
            ('SKIP', r'[ \t]+'),
            ('OPERATOR', r'<>|<=|>=|:=|[+\-*/^=&<>\(\)\.,:]'), # Added : to operator
            ('IDENTIFIER', r'[a-zA-Z_]\w*'), 
            ('MISMATCH', r'.'),
        ]
        
        # Compile regex
        self.master_pat = re.compile('|'.join('(?P<%s>%s)' % pair for pair in self.token_specs), re.IGNORECASE)

    def tokenize(self):
        for mo in self.master_pat.finditer(self.code):
            kind = mo.lastgroup
            value = mo.group()
            
            if kind == 'LINE_CONTINUATION':
                self.line += 1
                self.column = 1
                continue # Skip it entirely
            elif kind == 'NEWLINE':
                self.line += 1
                self.column = 1
                yield Token(kind, '\n', self.line, self.column)
                continue
            elif kind == 'SKIP':
                self.column += len(value)
                continue
            elif kind == 'MISMATCH':
                self.column += len(value)
                # yield Token('UNKNOWN', value, self.line, self.column)
                continue
            
            yield Token(kind, value, self.line, self.column)
            self.column += len(value)

        yield Token('EOF', '', self.line, self.column)
