import re


# VBA accepts a wide variety of date / time literal formats and
# Python's `datetime.strptime` is too strict for some (e.g. `1/1/100`
# fails because `%Y` insists on 4 digits). We validate structurally
# instead: split off an optional time component, check the date
# component against a small set of patterns, and range-check each
# field.

_DATE_PATTERNS = (
    # m/d/y or m/d/yyyy (US, the VBA default)
    re.compile(r"^\s*(?P<m>\d{1,2})\s*/\s*(?P<d>\d{1,2})\s*/\s*(?P<y>\d{1,4})\s*$"),
    # yyyy-mm-dd (ISO)
    re.compile(r"^\s*(?P<y>\d{1,4})\s*-\s*(?P<m>\d{1,2})\s*-\s*(?P<d>\d{1,2})\s*$"),
    # d-mmm-yyyy / d-mmm-yy (e.g. 1-Jan-2020)
    re.compile(
        r"^\s*(?P<d>\d{1,2})\s*-\s*(?P<mname>"
        r"jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|"
        r"january|february|march|april|may|june|july|august|"
        r"september|october|november|december"
        r")\s*-\s*(?P<y>\d{1,4})\s*$",
        re.IGNORECASE,
    ),
    # MMMM d, yyyy / MMMM d yyyy
    re.compile(
        r"^\s*(?P<mname>"
        r"jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|"
        r"january|february|march|april|may|june|july|august|"
        r"september|october|november|december"
        r")\s+(?P<d>\d{1,2})\s*,?\s+(?P<y>\d{1,4})\s*$",
        re.IGNORECASE,
    ),
)

_TIME_PATTERN = re.compile(
    r"^\s*(?P<h>\d{1,2})(?:\s*:\s*(?P<mi>\d{1,2}))?"
    r"(?:\s*:\s*(?P<s>\d{1,2}))?"
    r"(?:\s*(?P<ampm>am|pm))?\s*$",
    re.IGNORECASE,
)

_DAYS_IN_MONTH = (31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)


def _is_valid_date_part(m, d, y):
    if not (1 <= m <= 12):
        return False
    if not (1 <= y <= 9999):
        return False
    if not (1 <= d <= _DAYS_IN_MONTH[m - 1]):
        return False
    return True


def _is_valid_time_part(content):
    m = _TIME_PATTERN.match(content)
    if not m:
        return False
    h = int(m.group("h"))
    mi = int(m.group("mi") or 0)
    s = int(m.group("s") or 0)
    ampm = m.group("ampm")
    if ampm:
        # 12-hour clock: hour must be 1..12
        if not (1 <= h <= 12):
            return False
    else:
        if not (0 <= h <= 23):
            return False
    return 0 <= mi <= 59 and 0 <= s <= 59


def _is_valid_vba_date_literal(content):
    """Validate the text between the surrounding `#` of a VBA date literal.

    Accepts date-only, time-only, or date + time forms. Range-checks the
    individual components rather than relying on Python `strptime`,
    which is stricter than VBA (e.g. it rejects 3-digit years).
    """
    if content is None:
        return False
    s = content.strip()
    if not s:
        return False

    # Try splitting "<date> <time>" — the date contains `/` or `-` and the
    # time contains `:`. If both are present we test each side.
    parts = s.split()
    if len(parts) >= 2:
        # Find the split: contiguous date tokens up to where a `:` appears
        for i in range(len(parts) - 1, 0, -1):
            left = " ".join(parts[:i])
            right = " ".join(parts[i:])
            if ":" in right and (":" not in left):
                if _match_date(left) and _is_valid_time_part(right):
                    return True
        # AM/PM at the end
        if parts[-1].lower() in ("am", "pm"):
            left = " ".join(parts[:-2]) if len(parts) >= 2 else ""
            time_part = " ".join(parts[-2:]) if len(parts) >= 2 else ""
            if left and _match_date(left) and _is_valid_time_part(time_part):
                return True

    # Date only?
    if _match_date(s):
        return True
    # Time only?
    if _is_valid_time_part(s):
        return True
    return False


def _match_date(s):
    for pat in _DATE_PATTERNS:
        m = pat.match(s)
        if not m:
            continue
        gd = m.groupdict()
        try:
            day = int(gd["d"])
            year = int(gd["y"])
            if "m" in gd and gd.get("m"):
                month = int(gd["m"])
            else:
                month = _MONTH_NAME_TO_NUM[gd["mname"].lower()[:3]]
        except (KeyError, ValueError):
            continue
        if _is_valid_date_part(month, day, year):
            return True
    return False


_MONTH_NAME_TO_NUM = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


class Token:
    def __init__(self, type, value, line, column):
        self.type = type
        self.value = value
        self.line = line
        self.column = column

    def __repr__(self):
        return f"Token({self.type}, {repr(self.value)}, Line:{self.line})"

class LexerError:
    """Reported when the lexer encounters a character it cannot tokenize
    or a malformed literal (date, …)."""
    def __init__(self, char, line, column):
        self.char = char
        self.line = line
        self.column = column
        self.rule_id = "VBA_LEX001"
        self.message = (
            f"Lexer error: unexpected character {char!r} at line {line}, "
            f"column {column}."
        )

    def to_dict(self, filename=""):
        return {
            "file": filename,
            "line": self.line,
            "column": self.column,
            "rule_id": self.rule_id,
            "severity": "error",
            "message": self.message,
        }


class Lexer:
    def __init__(self, code):
        self.code = code
        self.pos = 0
        self.line = 1
        self.column = 1
        self.errors = []

        # Regex patterns
        # Identifiers may carry the legacy String type-suffix `$` directly
        # appended (`Mid$`, `Left$`, `Trim$`, `Format$`, …). Bracket-quoted
        # identifiers (`[A1]`, `[Sheet1!A1]`) are VBA's foreign-name escape
        # used heavily in Excel/host integration.
        self.token_specs = [
            ('COMMENT', r"'.*"),
            ('STRING', r'"(""|[^"])*"'),
            # DATELITERAL must come before PREPROCESSOR — both start with `#`
            # and the regex engine takes the first match in the alternation,
            # so PREPROCESSOR's `#[a-zA-Z_]\w*` would otherwise eat
            # `#January` from `#January 1, 2020#` and leave a stray `#`.
            ('DATELITERAL', r'\#[^#\r\n]+\#'),
            # VBA file-number argument used by I/O statements:
            #   Open path For Binary As #1
            #   Print #1, "x" / Put #1, , buf / Close #1
            # Numeric file-numbers start with `#<digit>+`; lexically
            # absent variants (`#fileVar`) already match PREPROCESSOR.
            ('FILENUMBER', r'#\d+'),
            ('PREPROCESSOR', r'#[a-zA-Z_]\w*'),
            # Numeric literals may carry a trailing legacy type-suffix:
            # & Long, % Integer, # Double, ! Single, @ Currency, $ String
            # (rarely on numeric, but harmless to allow).
            ('HEX', r'&H[0-9A-Fa-f]+[&%@!#]?'),
            ('OCTAL', r'&O[0-7]+[&%@!#]?'),
            ('FLOAT', r'(?:(?:\d+\.\d*|\.\d+|\d+)[eEdD][+\-]?\d+|\d+\.\d+)[#!@]?'),
            ('INTEGER', r'\d+[&%@!#]?'),
            # Line continuation — `_` must be preceded by whitespace, but
            # VBA tolerates trailing whitespace (and an inline `'` comment
            # is technically permitted before the newline; we keep it
            # simple and only swallow whitespace).
            ('LINE_CONTINUATION', r'[ \t]+_[ \t]*(\r\n|\n)'),
            ('NEWLINE', r'(\r\n|\n)'), # Removed : from newline
            ('SKIP', r'[ \t]+'),
            ('OPERATOR', r'<>|<=|>=|:=|[+\-*/^=&<>\(\)\.,:\\!]'), # Added : \ ! to operator
            ('BRACKET_IDENTIFIER', r'\[[^\]\r\n]*\]'),
            # Identifier may carry a legacy single-character type suffix:
            # $ → String, % → Integer, @ → Currency.
            # &, !, # are already used as operators / preprocessor / date
            # markers and stay tokenised separately to keep disambiguation
            # simple — the analyzer's _normalize_identifier is permissive
            # about the suffixes it strips.
            ('IDENTIFIER', r'[a-zA-Z_]\w*[$%@]?'),
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
                # Don't drop silently: capture so callers can surface the error
                # instead of silently producing a garbage token stream.
                self.errors.append(LexerError(value, self.line, self.column))
                self.column += len(value)
                continue
            elif kind == 'DATELITERAL':
                # Strip surrounding `#` and validate the contents.
                inner = value[1:-1] if len(value) >= 2 else value
                if not _is_valid_vba_date_literal(inner):
                    err = LexerError(value, self.line, self.column)
                    err.message = (
                        f"Invalid date literal {value!r} at line {self.line}, "
                        f"column {self.column}: not a recognised VBA date / time "
                        f"format."
                    )
                    err.rule_id = "VBA_LEX002"
                    self.errors.append(err)

            yield Token(kind, value, self.line, self.column)
            self.column += len(value)

        yield Token('EOF', '', self.line, self.column)
