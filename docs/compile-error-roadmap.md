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
| 3  | Argumentlisten-Wohlgeformtheit | VBA130–VBA136 | ⬜ geplant |
| 4  | Benannte Argumente an Aufrufstellen | VBA140–VBA142 | ⬜ geplant |
| 5  | Prozedur-/Property-Konsistenz | VBA150–VBA154 | ⬜ geplant |
| 6  | Zuweisungs-/Ausdrucks-Semantik | VBA160–VBA166 | ⬜ geplant |
| 7  | Schleifen-/Label-Kontext | VBA170–VBA175 | ⬜ geplant |
| 8  | Statement-Platzierung / Def / Option | VBA180–VBA185 | ⬜ geplant |
| 9  | Bedingte Kompilierung (`#If`) | VBA190–VBA192 | ⬜ geplant |
| 10 | Lexer-Ergänzungen | VBA_LEX003–006 | ⬜ geplant |

Regeln gesamt: **41 → 51** (nach Welle 1).

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

## Nächste Schritte

**Welle 2 (Deklarations-Typauflösung, VBA120–VBA123)** ist der nächste
höchste Hebel: `User-defined type not defined`, Forward-Reference, leere
`Type`/`Enum`. Höheres Falsch-Positiv-Risiko → neuer Validator
`_validate_declared_types(mod)` nur aktiv bei geladenem Host-/Std-Modell.
Details siehe Analyse-Doku, Teil 2.
