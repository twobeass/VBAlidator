# Analyse-Report: Nicht erkannte VBA-Compile-Fehler & vollständige Implementierungs-Roadmap

## Kontext

VBAlidator erkennt aktuell **41 Regeln** (`src/rules.py`). Ziel: anhand des
**autoritativen Microsoft-Learn-Fehlerkatalogs** (`MicrosoftDocs/VBA-Docs`,
`Language/Reference/error-messages` — ~400 Meldungen) maximale Vollständigkeit
herstellen: welche **statisch aus dem Quelltext erkennbaren** Compile-Fehler
fehlen und wie sie zu implementieren sind. Der Anwender will „nur Report" +
„im besten Fall alle Fehler erkannt" → dieses Dokument ist Report + komplette
Umsetzungs-Roadmap.

### Scope-Abgrenzung (was ein statischer Quelltext-Validator leisten kann)
Aus dem MS-Katalog **ausgeschlossen** (nicht statisch entscheidbar bzw. nicht
Quelltext-Sache):
- Runtime-Trappable-Errors mit Code, z. B. *Division by zero (11)*, *Overflow
  (6)*, *Subscript out of range (9)*, *Object variable not set (91)*,
  *Type mismatch (13)* zur Laufzeit, *File not found (53)* etc.
- IDE-/Lade-/Projekt-/Referenz-/Drucker-/ActiveX-/Registry-Fehler, *Line
  'item1': …*-Formulardatei-Meldungen, OLE/COM-Registrierung, Speicher.

**In-Scope** sind die reinen Compile-Time-/Syntax-/Semantik-Meldungen unten.

### Architektur (für die „Wie"-Spalte, verifiziert)
- Parser `src/parser.py`: rekursiver Descent. Blockparser rufen
  `parse_block(end_markers=[...])` + danach `self.consume(...)`. **`consume()`
  (parser.py:200) gibt bei fehlendem Token still `False` zurück, ohne Fehler.**
  Syntaxfehler über `_record_syntax_error(msg, line, rule_id)` (parser.py:181).
  Inverse Terminatoren werden bereits generisch erkannt (parse_block:714/723).
- Analyzer `src/analyzer.py`: Validatoren `_validate_xxx(mod)` hängen Dicts an
  `self.errors`. Vorlagen: `_validate_enum_uniqueness` (362),
  `_validate_property_arity` (392), `_validate_ptrsafe_declares`,
  `_validate_jump_target` (Label-Registry). Typ-Helfer:
  `_is_clearly_scalar` / `_is_clearly_object`. Discovery `pass1_discovery`
  (~110–155) registriert UDT/Enum/Vars; `mod.types` = UDT+Enum;
  Host-/Built-in-Typen aus dem geladenen Modell.
- Präprozessor `src/preprocessor.py`: `#If`-Auswertung (für Wave 9).
- Signatur-Daten: `proc.args` = `VariableNode` mit
  `is_optional`/`is_paramarray`/`mechanism`/`type_name`.
- Neue Regel = (1) `_RULES`-Eintrag in `src/rules.py`, (2) `rule_id` beim
  `append`, (3) `python tools/generate_rule_docs.py`, (4) Fixture unter
  `tests/samples/compile_errors/<kat>/` (auto-erfasst von
  `tests/test_compile_error_samples.py`) + Kategorie-Keywords dort, ggf.
  Direkttest. `tests/test_rules_registry.py` erzwingt Registry-Abdeckung.
  `tests/test_awesome_vba_regression.py` = Falsch-Positiv-Wächter (Ziel 0 FP).

---

## Teil 1 — Coverage des statisch erkennbaren Compile-Fehler-Katalogs

Legende: ✅ erkannt · 🟡 teilweise/generisch · ❌ nicht erkannt

### A. Block-/Terminator-Paarung
| MS-Meldung | Status | Regel/Anmerkung |
|---|---|---|
| Block If without End If / You must terminate the If block with an End If | ❌ | parse_if_stmt läuft bis EOF |
| End If without block If | 🟡 | generisch (parse_block:723) |
| For without Next | ❌ | parse_for:969 |
| Next without For | 🟡 | generisch (714) |
| Do without Loop | ❌ | parse_do:1004 |
| Loop without Do | 🟡 | generisch |
| While without Wend | ❌ | parse_while:795 |
| Wend without While | 🟡 | generisch |
| Select Case without End Select | ❌ | parse_select |
| End Select without Select Case | 🟡 | generisch |
| End With without With / Expected End With | 🟡/❌ | parse_with:816 |
| Case without Select Case | ❌ | |
| Case Else outside Select Case | ❌ | |
| Statements or labels invalid between Select Case and first Case | ❌ | |
| Expected End Sub/Function/Property | 🟡 | überlappt VBA350 |

### B. Deklaration & Typen
| MS-Meldung | Status | Regel |
|---|---|---|
| User-defined type not defined | ❌ | Typen werden nie aufgelöst |
| Forward reference to user-defined type | ❌ | |
| User-defined type without members not allowed | ❌ | |
| Empty Enum type not allowed | ❌ | |
| Duplicate declaration in current scope | ✅ | VBA003 |
| Duplicate definition / Duplicate procedure name | ❌ | s. Ambiguous name |
| Variable not defined | ✅ | VBA001 |
| Assignment to constant not permitted | ❌ | |
| Variable required. Can't assign to this expression | ❌ | |
| Array already dimensioned | 🟡 | VBA103 |
| Invalid ReDim / Can't assign to an array | ❌/🟡 | |
| Constant expression required | ✅ | VBA230/231 |
| Invalid data type for constant | ❌ | |
| Fixed-length strings nur Modulebene | ✅ | VBA250 |
| Invalid length for fixed-length string | ❌ | |
| Type-declaration character not allowed/required/mismatch | ❌ | |
| Identifier too long / Line too long / Too many line continuations | ❌ | |
| Deftype must precede declarations / Duplicate Deftype | ❌ | |
| Duplicate Option statement | ❌ | |

### C. Prozeduren, Parameter, Properties
| MS-Meldung | Status | Regel |
|---|---|---|
| Wrong number of arguments / Argument not optional | ✅ | VBA006 |
| ByRef argument type mismatch | ✅ | VBA007 |
| Argument required for property Let or Set | ✅ | VBA221 |
| Exit Function/Property/Sub im falschen Prozedur-Typ | ✅ | VBA008 |
| Procedure terminator mismatch | ✅ | VBA350 |
| Ambiguous name detected / A procedure of that name already exists | ❌ | nur Var/Const (VBA003) |
| Only comments may appear after End Sub/Function/Property | ❌ | |
| Definitions of property procedures … inconsistent | 🟡 | VBA222 nur Arity, nicht Typ |
| Procedure declaration does not match event/proc with same name | 🟡 | VBA330/340 nur Existenz |
| ParamArray must be declared as an array of variant | ❌ | |
| Can't have paramarrays with optional arguments | ❌ | |
| Array argument must be ByRef | ❌ | |
| Invalid ParamArray use | ❌ | |
| Optional argument must be variant / Invalid optional parameter type | ❌ | |
| Named argument not found | ❌ | |
| Named argument already specified | ❌ | |
| Named arguments not allowed / ParamArray-Proc mit benannten Args | ❌ | |
| Procedure type mismatch | ❌ | |

### D. Kontrollfluss, Schleifen, Labels
| MS-Meldung | Status | Regel |
|---|---|---|
| Label not defined | ✅ | VBA201 |
| Unreachable code (Warnung) | ✅ | VBA009 |
| Exit For not within For...Next | ❌ | |
| Exit Do not within Do...Loop | ❌ | |
| Invalid Next control variable reference | ❌ | `Next j` ≠ Schleifenvar |
| For control variable already in use | ❌ | verschachtelt gleiche Var |
| For Each control variable must be variant or object | ❌ | |
| Duplicate label | ❌ | Label-Registry vorhanden |

### E. Member-Zugriff / Ausdrücke / Objekte
| MS-Meldung | Status | Regel |
|---|---|---|
| Method or data member not found | ✅ | VBA002/VBA260 |
| Expected array | ✅ | VBA005 |
| Invalid `.member` ohne With | ✅ | VBA004 |
| Object module must implement all procedures in interface | ✅ | VBA330 |
| Object/Scalar-Zuweisung (Set) | ✅ | VBA210/211 |
| Invalid use of Me keyword | ❌ | |
| Invalid use of New keyword | ❌ | |
| Invalid use of AddressOf operator | ❌ | |
| Function call on LHS of assignment must return variant/object | ❌ | |
| Wrong number of dimensions / Too many dimensions | ❌ | |
| With object must be user-defined type, object, or variant | ❌ | |
| Expected procedure/variable, not module/project/Enum | 🟡 | tlw. VBA005 |

### F. Statement-Platzierung / Scope
| MS-Meldung | Status | Regel |
|---|---|---|
| Invalid inside/outside procedure | ✅ | VBA360/361 |
| Statement invalid inside/outside Type block | 🟡 | überlappt VBA360 |
| Invalid inside Enum / Invalid outside Enum | ❌ | |
| Must be first statement on the line | ❌ | |
| Qualified name disallowed in module scope | ❌ | |

### G. Bedingte Kompilierung (#If)
| MS-Meldung | Status | Regel |
|---|---|---|
| #Else / #ElseIf / #End If without matching #If | ❌ | preprocessor |
| Missing #End If | ❌ | |
| Invalid syntax for conditional compiler constant declarations | ❌ | |

### H. Lexer / Platform / Operatoren
| MS-Meldung | Status | Regel |
|---|---|---|
| Invalid character | ✅ | VBA_LEX001 |
| Invalid date literal | ✅ | VBA_LEX002 |
| Syntax error | ✅ | VBA010 |
| Arithmetik String↔Numerisch | ✅ | VBA240 |
| Declare ohne PtrSafe (64-bit) / Code must be updated for 64-bit | ✅ | VBA300 |
| Duplicate Enum member | ✅ | VBA310 |
| RaiseEvent ohne Event / falsche Arity | ✅ | VBA340/341 |
| Missing end bracket (unklammerter String/`[name]`) | ❌ | lexer |
| Invalid event name / Event not found | 🟡 | VBA340 |

**Fazit:** Der semantische Kern ist sehr gut abgedeckt. Die mit Abstand
wirkungsvollsten Lücken sind **Block-Terminator-Paarung (A)** und
**Deklarations-Typauflösung (B)**. Daneben existiert eine lange Reihe gut
implementierbarer, rein syntaktischer Regeln (C/D/F/G/H).

---

## Teil 2 — Implementierungs-Roadmap (priorisiert, mit Regel-IDs)

Reihenfolge nach Nutzen × geringes Falsch-Positiv-Risiko. Wellen 1, 3, 5–9
sind überwiegend rein **syntaktisch** (kein Typsystem nötig → geringes
FP-Risiko). Wellen 2 und 4 brauchen Symbol-/Typkontext (höheres FP-Risiko →
nur bei geladenem Modell aktiv).

### Welle 1 — Block-Terminator-Paarung *(rein Parser, höchster Nutzen)*
Neue IDs **VBA110–VBA119**, Phase 1.5.
- VBA110 For without Next · VBA111 Do without Loop · VBA112 While without Wend
  · VBA113 Expected End With · VBA114 Select Case without End Select
  · VBA115 Block If without End If.
- **Wie:** in `parse_for/parse_do/parse_while/parse_with/parse_select/parse_if_stmt`
  den Rückgabewert von `self.consume('IDENTIFIER','<Terminator>')` prüfen; bei
  `False` `_record_syntax_error("Compile error: For without Next", line, "VBA110")`.
  Beispiel parse_for:969:
  ```python
  if not self.consume('IDENTIFIER', 'Next'):
      self._record_syntax_error("Compile error: For without Next",
                                line=line, rule_id="VBA110")
  ```
- VBA116–VBA119: die bereits generisch erkannten Inversen
  (parse_block:714/723) auf dedizierte IDs heben: Next/Loop/Wend without …,
  End If/End Select/End With without opener, **Case without Select Case**,
  **Case Else outside Select Case**, **Statements/labels between Select Case
  and first Case** (in `parse_select` prüfen).

### Welle 2 — Deklarations-Typauflösung *(Analyzer, nur mit Modell)*
Neue IDs **VBA120–VBA123**, Phase 1.6. Neuer Validator
`_validate_declared_types(mod)`.
- VBA120 User-defined type not defined: über alle `type_name` (Modul-Vars,
  Locals, Args, UDT-Member, Return-Typen) prüfen gegen: Built-ins (vgl.
  `_is_clearly_scalar/_object`), `mod.types`, Projekt-`.cls`-Namen,
  Host-/Std-Modell. Punktierte Typen am letzten Segment. **FP-Schutz:** nur
  melden wenn Modell geladen; sonst Severity `warning` oder per Flag.
- VBA121 Forward reference to user-defined type (UDT-Member referenziert
  später deklarierten Typ — Reihenfolge in `mod.types`).
- VBA122 User-defined type without members not allowed (leerer `Type … End
  Type`).
- VBA123 Empty Enum type not allowed.

### Welle 3 — Argumentlisten-Wohlgeformtheit *(Analyzer, rein syntaktisch)*
Neue IDs **VBA130–VBA136**, Phase 2.10. Validator `_validate_param_lists(mod)`
über `proc.args`.
- VBA130 Optional vor Pflichtparameter · VBA131 ParamArray nicht letzter
  · VBA132 mehrere ParamArray · VBA133 ParamArray + Optional verboten
  · VBA134 ParamArray muss `() As Variant` sein · VBA135 Invalid optional
  parameter type / Optional argument must be variant (kontextabhängig)
  · VBA136 Array argument must be ByRef.

### Welle 4 — Benannte Argumente an Aufrufstellen *(Analyzer, Signaturkontext)*
Neue IDs **VBA140–VBA142**. Erweitert die bestehende Call-Validierung
(analyzer.py:~1982–2040).
- VBA140 Named argument not found · VBA141 Named argument already specified
  · VBA142 ParamArray-Prozedur mit benannten Args. FP-Schutz wie Welle 2.

### Welle 5 — Prozedur-/Property-Konsistenz *(Analyzer/Parser)*
Neue IDs **VBA150–VBA154**.
- VBA150 Ambiguous name detected / Duplicate procedure name: Validator über
  `mod.procedures` nach Name gruppiert (Property Get/Let/Set ausnehmen).
- VBA151 Only comments may appear after End Sub/Function/Property: in
  `procedures_parse` nach Terminator (parser.py:669) bis zur nächsten
  Prozedur nur Kommentare/Leerzeilen zulassen.
- VBA152 Property-Prozeduren inkonsistent: `_validate_property_arity` (392)
  erweitern — Get-Rückgabetyp ↔ Let/Set-Wertparametertyp.
- VBA153 Signatur stimmt nicht mit Interface/Event überein: VBA330/340 um
  Parametertyp-Vergleich erweitern.
- VBA154 Procedure type mismatch (AddressOf-Ziel / Event-Handler-Signatur).

### Welle 6 — Zuweisungs-/Ausdrucks-Semantik *(Analyzer)*
Neue IDs **VBA160–VBA166**.
- VBA160 Assignment to constant not permitted / Variable required (LHS ist
  Const/Enum-Member/Literal — im Set/Let-Pfad analyzer.py:~691).
- VBA161 Can't assign to an array · VBA162 Function call on LHS must return
  variant/object · VBA163 Wrong/Too many dimensions · VBA164 Invalid use of
  Me · VBA165 Invalid use of New · VBA166 Invalid use of AddressOf.

### Welle 7 — Schleifen-/Label-Kontext *(Analyzer)*
Neue IDs **VBA170–VBA175**.
- VBA170 Exit For not within For...Next · VBA171 Exit Do not within Do...Loop
  (analog zu VBA008, Kontext-Stack in `analyze_block`).
- VBA172 Invalid Next control variable reference (`parse_for`:971 muss den
  optionalen Next-Namen am `ForNode` speichern; Analyzer vergleicht mit
  `var_token`).
- VBA173 For control variable already in use (verschachtelte For mit gleicher
  Variable) · VBA174 For Each control variable must be variant or object
  · VBA175 Duplicate label (Label-Registry um Duplikat-Erkennung erweitern).

### Welle 8 — Statement-Platzierung / Def / Option *(Parser)*
Neue IDs **VBA180–VBA185**.
- VBA180 Invalid inside/outside Enum · VBA181 Statement invalid inside/outside
  Type block (VBA360 erweitern) · VBA182 Deftype must precede declarations
  · VBA183 Duplicate Deftype · VBA184 Duplicate Option statement
  · VBA185 Must be first statement on the line.

### Welle 9 — Bedingte Kompilierung *(Preprocessor)*
Neue IDs **VBA190–VBA192** in `src/preprocessor.py`.
- VBA190 #Else/#ElseIf/#End If without matching #If (Stack-Balance)
  · VBA191 Missing #End If · VBA192 Invalid conditional-compiler constant
  declaration syntax.

### Welle 10 — Lexer-Ergänzungen *(Lexer)*
Neue IDs **VBA_LEX003–VBA_LEX006** in `src/lexer.py`.
- VBA_LEX003 Type-declaration character mismatch/not allowed/required
  · VBA_LEX004 Missing end bracket (unklammerter String / `[name]`)
  · VBA_LEX005 Identifier too long (>255) · VBA_LEX006 Line too long / Too
  many line continuations.

---

## Teil 3 — Verifikation (für jede künftige Welle)
Reiner Report → keine Code-Änderung. Pro umgesetzter Regel:
1. Fixture(s) unter `tests/samples/compile_errors/<kat>/` ablegen + Kategorie-
   Keywords in `tests/test_compile_error_samples.py` ergänzen.
2. `pytest -q` (Linux genügt; Roundtrip optional).
3. `python tools/generate_rule_docs.py` → `docs/rules/`; Registry-Test
   (`tests/test_rules_registry.py`) erzwingt `_RULES`-Eintrag pro `rule_id`.
4. `tests/test_awesome_vba_regression.py` als Falsch-Positiv-Wächter (Ziel 0
   FP) — besonders Wellen 2 & 4.

## Quellen
- MicrosoftDocs/VBA-Docs — `Language/Reference/error-messages` (vollständige
  Meldungsliste, ~400 Einträge; via raw.githubusercontent.com bezogen).
- Microsoft Learn — *Block If without End If*; *User-defined type not defined*.
- software-solutions-online.com — *Next Without For*, *End If without block If*.
