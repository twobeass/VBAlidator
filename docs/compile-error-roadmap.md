# Compile-Error Coverage — Roadmap & Live-Status

Begleitdokument zu [`compile-error-analysis.md`](compile-error-analysis.md).
Hier wird der **aktuelle Umsetzungsstand** der dort beschriebenen Wellen
gepflegt. Quelle der Lücken: autoritativer Microsoft-Katalog
(`MicrosoftDocs/VBA-Docs`, `Language/Reference/error-messages`, ~400 Meldungen),
gefiltert auf statisch aus dem Quelltext erkennbare Compile-Fehler.

**Statuslegende:** ✅ fertig · 🟡 in Arbeit · ⬜ geplant

## Gesamtfortschritt

| Welle | Thema | Regel-IDs | Status |
|-------|-------|-----------|--------|
| 1  | Block-Terminator-Paarung | VBA110–VBA119 | ✅ fertig |
| 2  | Deklarations-Typauflösung | VBA120–VBA123 | ⬜ geplant |
| 3  | Argumentlisten-Wohlgeformtheit | VBA130–VBA136 | ✅ fertig |
| 4  | Benannte Argumente an Aufrufstellen | VBA140–VBA142 | ⬜ geplant |
| 5  | Prozedur-/Property-Konsistenz | VBA150–VBA154 | ⬜ geplant |
| 6  | Zuweisungs-/Ausdrucks-Semantik | VBA160–VBA166 | ⬜ geplant |
| 7  | Schleifen-/Label-Kontext | VBA170–VBA175 | ✅ fertig |
| 8  | Statement-Platzierung / Def / Option | VBA180–VBA184 (VBA185 zurückgestellt) | ✅ fertig |
| 9  | Bedingte Kompilierung (`#If`) | VBA190–VBA192 | ✅ fertig |
| 10 | Lexer-Ergänzungen | VBA_LEX004–006 (VBA_LEX003 zurückgestellt) | ✅ fertig |

Regeln gesamt: **41 → 74** (nach Welle 10).

---

## Welle 1 — Block-Terminator-Paarung ✅

Abgeschlossen. Jeder Block-Opener wird jetzt auf seinen passenden Terminator
geprüft; ein verwaister Terminator ohne Opener wird ebenfalls gemeldet.

**Neue Regeln (Phase 1.5):**

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA110 | For without Next | `For`/`For Each` ohne `Next` |
| VBA111 | Do without Loop | `Do` ohne `Loop` |
| VBA112 | While without Wend | `While` ohne `Wend` |
| VBA113 | Expected End With | `With` ohne `End With` |
| VBA114 | Select Case without End Select | `Select Case` ohne `End Select` |
| VBA115 | Block If without End If | mehrzeiliges `If … Then` ohne `End If` |
| VBA116 | Next without For | verwaistes `Next` |
| VBA117 | Loop without Do | verwaistes `Loop` |
| VBA118 | Wend without While | verwaistes `Wend` |
| VBA119 | Block terminator without matching opener | `Else`/`ElseIf`/`End If`/`End Select`/`End With` ohne Opener |

**Umsetzung:**
- `src/parser.py`: `parse_for`/`parse_do`/`parse_while`/`parse_with`/
  `parse_select`/`parse_if_stmt` prüfen jetzt den Rückgabewert von
  `consume(...)`; bei fehlendem Terminator → `_record_syntax_error(..., rule_id=…)`.
  Block-If verfolgt `end_if_seen`. `parse_block` mappt die inversen
  Terminatoren auf VBA116–VBA119.
- `src/rules.py`: Registry-Einträge VBA110–VBA119 (Abschnitt „Phase 1.5").
- `docs/rules/VBA11x.md`: via `tools/generate_rule_docs.py` generiert.

**Tests (alle grün):**
- `tests/test_phase15_block_terminators.py`: je Regel ein Direkttest +
  „balanced blocks bleiben sauber".
- `tests/samples/compile_errors/block_terminator/`: 7 Fixtures,
  Kategorie-Keywords in `tests/test_compile_error_samples.py`.
- Falsch-Positiv-Wächter `tests/test_awesome_vba_regression.py`: unverändert
  grün (Terminator-Checks feuern nur bei echtem Fehlen).

**Bekannte Einschränkung (Refinement für später):** Bei einem fehlenden
Terminator wird zusätzlich zur primären Meldung (z. B. VBA110) gelegentlich
ein sekundäres VBA119 für das verschluckte `End Sub` erzeugt. Beide zeigen auf
dieselbe Fehlerregion; die primäre Meldung ist korrekt. Eine Entkopplung
(Proc-Terminatoren nicht in `parse_block` konsumieren) ist als spätere
Verfeinerung notiert.

---

## Welle 3 — Argumentlisten-Wohlgeformtheit ✅

Abgeschlossen. Neuer Validator `_validate_param_lists(mod)` (analyzer.py),
rein syntaktisch über `proc.args`, aufgerufen in `pass2_resolution`.

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA130 | Required parameter after Optional | Pflichtparameter nach `Optional` |
| VBA131 | ParamArray is not the last parameter | `ParamArray` nicht als letzter |
| VBA132 | More than one ParamArray | mehr als ein `ParamArray` |
| VBA133 | ParamArray combined with Optional | `Optional` + `ParamArray` gemischt |
| VBA134 | ParamArray element type is not Variant | `ParamArray x() As <≠Variant>` |
| VBA136 | Array parameter passed ByVal | `ByVal arr() As …` |

**Wichtige FP-Korrektur:** Der Wert-Parameter (`RHS`) eines `Property
Let`/`Set` ist implizit Pflicht und darf nach Optional-Index-Parametern
stehen — der letzte Parameter von Let/Set ist daher von VBA130 ausgenommen
(sonst Falsch-Positive in realen Bibliotheken wie JSONBag/stdVBA).

**VBA135 (Invalid optional parameter type) zurückgestellt:** kontextabhängig,
kein eindeutiger Compile-Fehler → hohes FP-Risiko, bewusst nicht implementiert.

**Tests:** `tests/test_phase210_param_lists.py` (je Regel + „well-formed bleibt
sauber"), Fixtures in `tests/samples/compile_errors/param_list/`. FP-Wächter
grün.

---

## Welle 7 — Schleifen-/Label-Kontext ✅

Abgeschlossen. Loop-Kontext-Stacks (`self._loop_stack`, `self._for_var_stack`)
werden pro Prozedur in `analyze_procedure` zurückgesetzt und in `analyze_block`
beim Betreten von `ForNode`/`DoNode` gepusht/gepoppt. `parse_for` erfasst jetzt
die Variable des schließenden `Next` (inkl. `Next j, i`-Mehrfachschließung).

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA170 | Exit For not within For...Next | `Exit For` außerhalb einer For-Schleife |
| VBA171 | Exit Do not within Do...Loop | `Exit Do` außerhalb `Do` (auch in `While...Wend`) |
| VBA172 | Next control variable mismatch | `Next j` ≠ For-Zähler (außer `Next j, i`) |
| VBA173 | For control variable already in use | verschachtelte For mit gleichem Zähler |
| VBA174 | For Each control variable must be Variant/object | `For Each i` mit `i As Long` o. ä. |
| VBA175 | Duplicate label | Label zweimal in derselben Prozedur |

**Bekannte Einschränkung:** Die `Next j, i`-Mehrfachschließung wird vom Parser
nur am innersten Loop verarbeitet; die äußere Schleife kann dadurch ein
fehlerhaftes VBA110 erzeugen (selten in der Praxis, im Test-Korpus nicht
vorhanden). VBA172 ist daher bewusst auf die Einzel-`Next`-Form beschränkt.

**Tests:** `tests/test_phase211_loop_label.py` (12 Fälle inkl. Gegenbeispiele),
Fixtures in `tests/samples/compile_errors/loop_label/`. FP-Wächter grün.

---

## Welle 8 — Statement-Platzierung / Def / Option ✅

Abgeschlossen (parser-seitig).

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA180 | Statement invalid inside Enum block | fehlplatziertes Statement in `Enum` |
| VBA181 | Statement invalid inside Type block | fehlplatziertes Statement in `Type` |
| VBA182 | Deftype must precede declarations | `DefXxx` nach erster Deklaration |
| VBA183 | Duplicate Deftype letter range | überlappende `DefXxx`-Buchstabenbereiche |
| VBA184 | Duplicate Option statement | dieselbe `Option`-Anweisung doppelt |

**Wichtige FP-Korrektur (VBA180/181):** VBA erlaubt viele Keywords als
Feldnamen (`Type As Long`, `Name As String`, …). Die ursprüngliche
Keyword-Blacklist erzeugte 38 Falsch-Positive in stdVBA. Ersetzt durch einen
**Shape-Test**: eine Zeile, die mit zwei aufeinanderfolgenden Identifiern
beginnt und deren zweiter nicht `As` ist (z. B. `Dim x …`, `Sub Foo`), kann
keine Member-Deklaration sein → fehlplatziertes Statement. FP-Wächter grün.

**VBA185 (Must be first statement on the line) zurückgestellt:** kontext- und
formatabhängig, hohes FP-Risiko.

**Tests:** `tests/test_phase37_placement_def_option.py` (je Regel +
Gegenbeispiele), Fixtures in `tests/samples/compile_errors/placement_def_option/`.

---

## Welle 9 — Bedingte Kompilierung ✅

Abgeschlossen (preprocessor-seitig). `Preprocessor` sammelt jetzt Fehler
(`self.errors`) und bekommt einen `filename`; alle vier Aufrufstellen
(`api.py` ×2, `conftest.py` ×2) hängen `pp.errors` an `analyzer.errors`.

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA190 | Conditional directive without matching #If | `#Else`/`#ElseIf`/`#End If` ohne offenes `#If` (Stack-Unterlauf) |
| VBA191 | Missing #End If | Stack am Ende > Root → unterminiertes `#If` |
| VBA192 | Invalid #Const declaration | `#Const` ohne Name oder `= expression` |

**Tests:** `tests/test_phase38_conditional_compilation.py` (6 Fälle inkl.
Gegenbeispiele), Fixtures in `tests/samples/compile_errors/conditional_compilation/`.

---

## Welle 10 — Lexer-Ergänzungen ✅

Abgeschlossen (lexer-seitig).

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA_LEX004 | Unterminated string / missing end bracket | `"…` ohne `"`, `[…` ohne `]` |
| VBA_LEX005 | Identifier too long | Bezeichner > 255 Zeichen |
| VBA_LEX006 | Line too long | physische Zeile > 1023 Zeichen |

**VBA_LEX003 (Type-declaration character mismatch) zurückgestellt:** braucht
Parser-/Deklarationskontext (`Dim x% As String`), gehört nicht auf Lexer-Ebene
— bewusst nicht implementiert (FP-Risiko, falsche Schicht).

**Tests:** `tests/test_phase39_lexer_limits.py` (7 Fälle inkl. Gegenbeispiele),
Fixtures in `tests/samples/compile_errors/lexer_limits/`.

---

## Nächste Schritte

**Welle 2 (Deklarations-Typauflösung, VBA120–VBA123)** ist der nächste
höchste Hebel: `User-defined type not defined`, Forward-Reference, leere
`Type`/`Enum`. Höheres Falsch-Positiv-Risiko → neuer Validator
`_validate_declared_types(mod)` nur aktiv bei geladenem Host-/Std-Modell.
Details siehe Analyse-Doku, Teil 2.
