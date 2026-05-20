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
| 2  | Deklarations-Typauflösung | VBA120,122,123 (VBA121 zurückgestellt) | ✅ fertig |
| 3  | Argumentlisten-Wohlgeformtheit | VBA130–VBA136 | ✅ fertig |
| 4  | Benannte Argumente an Aufrufstellen | VBA140–VBA142 | ✅ fertig |
| 5  | Prozedur-/Property-Konsistenz | VBA150,152,153 (151/154 zurückgestellt) | ✅ fertig |
| 6  | Zuweisungs-/Ausdrucks-Semantik | VBA160,164,165 (161/162/163/166 zurückgestellt) | ✅ fertig |
| 7  | Schleifen-/Label-Kontext | VBA170–VBA175 | ✅ fertig |
| 8  | Statement-Platzierung / Def / Option | VBA180–VBA184 (VBA185 zurückgestellt) | ✅ fertig |
| 9  | Bedingte Kompilierung (`#If`) | VBA190–VBA192 | ✅ fertig |
| 10 | Lexer-Ergänzungen | VBA_LEX004–006 (VBA_LEX003 zurückgestellt) | ✅ fertig |

Regeln gesamt: **41 → 90** (10 Wellen + Maximum-Coverage-Nachschub).

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

**`Next j, i`-Mehrfachschließung:** im Review behoben. Der innerste `parse_for`
konsumiert die komplette Variablenliste und merkt sich über
`self._pending_next_closes`, wie viele *umschließende* Schleifen mitgeschlossen
wurden; `parse_block` beendet die betroffenen Frames sauber, sodass kein
fälschliches VBA110/VBA119 entsteht (Tests in
`test_phase15_block_terminators.py`). VBA172 bleibt bewusst auf die Einzel-
`Next`-Form beschränkt (Mehrfachform hat keinen eindeutigen Einzelbezug).

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

## Welle 5 — Prozedur-/Property-Konsistenz ✅

Abgeschlossen.

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA150 | Ambiguous name detected (duplicate procedure) | zwei Prozeduren gleichen Namens (Property Get/Let/Set ausgenommen) |
| VBA152 | Property Get/Let type inconsistent | Get-Rückgabetyp ≠ Let-Wertparametertyp (nur klar-skalare Typen) |
| VBA153 | Implements method signature mismatch | Implementierungsmethode mit falscher Parameteranzahl |

**Umsetzung:** neuer Validator `_validate_duplicate_procedures` (VBA150);
VBA152 in `_validate_property_arity` ergänzt (auf klar-skalare Typen
beschränkt → FP-sicher); VBA153 erweitert den bestehenden `_validate_implements`
um einen Parameteranzahl-Abgleich.

**VBA151 (Only comments after End Sub/Function/Property) zurückgestellt:**
überlappt mit VBA361 (executable at module level). **VBA154 (Procedure type
mismatch / AddressOf) zurückgestellt:** erfordert Signatur-Typsystem, hohe
Komplexität/FP.

**Tests:** `tests/test_phase212_proc_consistency.py` (7 Fälle inkl.
Mehrmodul-Implements), Fixtures in `tests/samples/compile_errors/proc_consistency/`.

---

## Welle 2 — Deklarations-Typauflösung ✅

Abgeschlossen. Neuer Validator `_validate_declared_types(mod)` + Resolver
`_is_known_type` / `_check_type_known`; lokale `Dim`-Typen zusätzlich in
`process_dim` geprüft.

| ID | Meldung | Severity | Auslöser |
|----|---------|----------|----------|
| VBA120 | User-defined type not defined | **warning** | unauflösbarer Typname |
| VBA122 | User-defined type without members | error | leeres `Type … End Type` |
| VBA123 | Empty Enum not allowed | error | leeres `Enum … End Enum` |

**Severity-Entscheidung (wichtig):** VBA120 ist **warning**, nicht error. Mit
einem prinzipiell unvollständigen Host-Modell lässt sich ein Tippfehler nicht
sicher von einem nicht-modellierten Host-Typ unterscheiden. VBA120 macht das
Risiko sichtbar, ohne die Compile-Safety zu blockieren — und der FP-Wächter
(zählt nur `error`) bleibt grün. VBA122/123 sind modellunabhängig und bleiben
harte Fehler.

**FP-Reduktion:** kuratierte Allowlist universeller VBA/stdole/MSForms-
Intrinsics (`OLE_COLOR`, `vbTriState`, `Form`, `Font`, `Picture`, …) +
großzügige Behandlung bibliotheksqualifizierter Namen (`Scripting.Dictionary`).

**VBA121 (Forward reference to UDT) zurückgestellt:** subtile Sonderfall-
Semantik, geringer Nutzen, FP-Risiko.

**Tests:** `tests/test_phase16_declared_types.py` (9 Fälle), Fixtures in
`tests/samples/compile_errors/declared_type/`. FP-Wächter grün.

---

## Welle 6 — Zuweisungs-/Ausdrucks-Semantik ✅

Abgeschlossen.

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA160 | Assignment to constant not permitted | Zuweisung an `Const`/Enum-Member |
| VBA164 | Invalid use of Me keyword | Zuweisung an `Me` |
| VBA165 | Invalid use of New keyword | `As New <primitive/Object/Variant>` |

**Umsetzung:** VBA160/164 in `_validate_set_vs_let` integriert; VBA165 in
`process_dim` (Flag `saw_new`). Nebenbei einen Parser-Quirk behoben: ein
modulweites bloßes `Const X = …` (ohne `Public`/`Private`) wurde mit
`is_const=False` registriert — jetzt korrekt als `Const` (Default-Scope
Private), wodurch auch VBA231/VBA160 dafür greifen.

**Zurückgestellt:** VBA161 (Can't assign to an array), VBA162 (Function call
on LHS), VBA163 (Wrong/Too many dimensions), VBA166 (Invalid use of AddressOf)
— erfordern ein Ausdruck-Typsystem, das der Analyzer nicht vollständig besitzt;
hohes FP-Risiko.

**Tests:** `tests/test_phase213_assignment_semantics.py` (6 Fälle), Fixtures in
`tests/samples/compile_errors/assignment_semantics/`. FP-Wächter grün.

---

## Welle 4 — Benannte Argumente an Aufrufstellen ✅

Abgeschlossen. Neuer Helfer `_validate_named_args`, aufgerufen aus
`validate_signature` (greift für geklammerte Aufrufe auflösbarer Prozeduren).

| ID | Meldung | Auslöser |
|----|---------|----------|
| VBA140 | Named argument not found | `name:=` ohne passenden Parameter (nur bei zuverlässig bekannten Quell-Signaturen) |
| VBA141 | Named argument specified more than once | doppeltes `name:=` im selben Aufruf (rein syntaktisch) |
| VBA142 | Named arguments with ParamArray | benannte Argumente bei `ParamArray`-Prozedur |

**FP-Schutz (im Review verschärft):** VBA140 feuert ausschließlich für
gescannte Quell-Prozeduren (`isinstance(extra, ProcedureNode)`). Host-Modell-
Parameternamen sind nur beratend (sie können den echten API-Namen abkürzen/
umbenennen, z. B. `Range.Find(What:=…)`) und lösen daher kein VBA140 aus.
VBA141 ist rein syntaktisch.

**Bekannte Einschränkung:** validiert wird nur die geklammerte Aufrufform
(`Foo(a:=1)`); die klammerlose Statement-Form (`Foo a:=1`) erreicht
`validate_signature` nicht und bleibt ungeprüft.

---

## Maximum-Coverage-Nachschub (nach dem Review)

Erneute kritische Durchsicht der zurückgestellten Regeln und des
Microsoft-Katalogs: alles, was sich FP-arm umsetzen lässt, wurde nachgezogen
(jeweils vorab gegen den realen `awesome_vba`-Korpus auf 0 Treffer geprüft):

| ID | Meldung | Schicht |
|----|---------|---------|
| VBA137 | User-defined type passed ByVal (`ByVal p As <UDT>`; Enums ausgenommen) | Analyzer |
| VBA151 | Code/Statement nach `End Sub/Function/Property` (nur Kommentare erlaubt) | Parser |
| VBA270 | Typkennzeichen ≠ `As`-Typ (`Dim x% As Long`; ersetzt das geplante VBA_LEX003) | Analyzer |
| VBA370 | Public Const/Array/Declare als Member eines Objektmoduls (Class/Form) | Analyzer |

## Status: abgeschlossen 🎉

Alle 10 Wellen + Maximum-Coverage-Nachschub. **41 → 90 Regeln.**

Verbleibende, bewusst **nicht** umgesetzte Regeln — jede würde ein vollständiges
Ausdruck-/Array-Typsystem erfordern oder ist nicht eindeutig statisch
entscheidbar (Umsetzung würde Falsch-Positive riskieren, was dem Zweck eines
Compile-Safety-Tools widerspricht):

| Regel | Thema | Grund der Nichtumsetzung |
|-------|-------|--------------------------|
| VBA121 | Forward reference to UDT | Sonderfall-Semantik, in der Praxis selten relevant |
| VBA135 | Invalid optional parameter type | kein eindeutiger Compile-Fehler |
| VBA154 | Procedure type mismatch (AddressOf) | Signatur-Typsystem für AddressOf-Ziele nötig |
| VBA161 | Can't assign to an array | Array-Kopie (`a = b`) ist gültiges VBA → FP-Risiko |
| VBA162 | Function call on LHS must return … | Rückgabetyp-Inferenz auf der linken Seite nötig |
| VBA163 | Too many / wrong number of dimensions | Array-Shape-Tracking nötig; >60 Dims praktisch irrelevant |
| VBA166 | Invalid use of AddressOf | Ziel kann in anderem Modul liegen → FP-Risiko |
| VBA185 | Must be first statement on the line | format-/kontextabhängig |

Der `awesome_vba`-Falsch-Positiv-Wächter ist über **alle** Wellen und den
Nachschub grün geblieben (395 Tests).
