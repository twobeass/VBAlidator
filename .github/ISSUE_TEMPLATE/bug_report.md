---
name: Bug Report (False Positive/Negative)
about: Report a case where vbalidator failed to correctly analyze VBA code.
title: "[BUG]: "
labels: bug, needs-triage
assignees: ''

---

## 1. Issue Type
- [ ] **False Positive** (The code is VALID, but vbalidator reported an error)
- [ ] **False Negative** (The code is INVALID, but vbalidator said it was fine)
- [ ] **Crash/Runtime Error** (The tool itself crashed)

## 2. The VBA Code Snippet
```vba
' Paste your code here
Sub Example()
    ' ...
End Sub
```

## 3. vbalidator Output
(Paste the tool's output here)

## 4. Expected Behavior

## 5. Environment (Optional)
 * OS:
 * vbalidator Version:

---

### Why this works for an AI Agent (Jules)

1.  **Classification (Section 1):** The AI immediately knows the goal.
    * If **False Positive**: The goal is to *loosen* constraints or add a missing definition.
    * If **False Negative**: The goal is to *tighten* logic or add a new rule.
2.  **Isolated Context (Section 2):** By forcing the user to put code in a triple-backtick block (` ```vba `), your agent can easily regex/extract the code snippet to run its own tests or reproduce the failure.
3.  **Ground Truth (Section 4):** This provides the "Label" for the training/fixing instance. It tells the agent *what* the logic should have been.