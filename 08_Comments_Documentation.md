# 08 - Comments & Documentation

## Objective

Comments and documentation exist to improve understanding where code alone is not sufficient.

However, well-written code should be self-explanatory. Comments are a supplement, not a substitute for clarity.

Bad comments explain **what the code already clearly says**. Good comments explain **why the code exists or why a decision was made**.

---

# Golden Rule

> If you need a comment to explain what the code does, the code is probably wrong.

---

# When to Use Comments

Use comments to explain:

- business rules
- non-obvious decisions
- external constraints (Excel, SAP, APIs, legacy systems)
- performance trade-offs
- known limitations or hacks
- historical context

---

# When NOT to Use Comments

Do NOT use comments for:

- obvious code behavior
- variable declarations
- step-by-step repetition of logic
- rewriting code in natural language

Bad:

```vb
' Set ws to worksheet
Set ws = Worksheets("Orders")
```

Good:

```vb
Set ws = Worksheets("Orders")
```

---

# Comment Style

Use clear, concise English sentences.

Start with a capital letter.

Do not over-comment.

Good:

```vb
' SAP requires leading zeros for material codes
articleCode = FormatArticle(articleCode)
```

Bad:

```vb
' we format the article code because sap needs it
```

---

# Business Rules Documentation

Always document non-obvious business logic.

Example:

```vb
' VAT is applied only if customer is in EU and order value > 1500
If isEUCustomer And orderTotal > 1500 Then
```

These rules are critical and must be explained.

---

# Workarounds and Hacks

Always document hacks or workarounds.

Example:

```vb
' Workaround: Excel API fails when range exceeds 10k rows
' Split export into chunks
Call ExportInChunks(ws)
```

---

# External System Constraints

Document constraints imposed by external systems.

Examples:

```vb
' SAP field length limitation: max 18 characters
articleCode = Left(articleCode, 18)
```

```vb
' Printer requires CRLF line endings
zpl = Replace(zpl, vbLf, vbCrLf)
```

---

# TODO Comments

Use TODOs sparingly and explicitly.

Format:

```vb
' TODO: Refactor this logic into BuildZplLabel
```

Never leave vague TODOs:

Bad:

```vb
' TODO fix this
```

Good:

```vb
' TODO: Handle multi-page label export for large orders
```

---

# FIXME Comments

Use for known bugs or temporary issues.

```vb
' FIXME: Crash occurs when worksheet is protected
```

All FIXMEs must be tracked externally if possible.

---

# Region Comments (Forbidden Style)

Do NOT use decorative separators:

Bad:

```vb
'==========================
' EXPORT SECTION
'==========================
```

Instead:

- split into smaller procedures
- use meaningful procedure names

---

# Inline Comments

Use inline comments only when necessary.

Good:

```vb
discount = price * 0.1 ' 10% promotional discount
```

Bad:

```vb
discount = price * 0.1 ' multiply price by 0.1
```

---

# Header Comments for Procedures

Only use header comments when logic is non-trivial.

Example:

```vb
' Builds ZPL label for SAP MB90 export
' Applies formatting rules required by Zebra printers
Public Function BuildZplLabel(ws As Worksheet, rowIndex As Long) As String
```

Do not restate obvious information.

---

# Self-Documenting Code First

Prefer meaningful naming over comments.

Bad:

```vb
' Get last row
lastRow = FindLastRow(ws)
```

Good:

```vb
lastRow = FindLastRow(ws)
```

Even better:

```vb
lastRow = GetLastDataRow(ws)
```

---

# Comment Density Rule

Aim for:

- 0–10% of lines commented
- higher only in complex business modules

If more comments are needed, refactor the code instead.

---

# Outdated Comments

Outdated comments are worse than no comments.

Always update comments when code changes.

---

# Disabled Code

Never leave commented-out code in production.

Bad:

```vb
' Call OldExportMethod(ws)
```

Remove it or use version control (Git).

---

# Debug Comments

Temporary debug comments are allowed during development:

```vb
' Debug: printing row index
Debug.Print rowIndex
```

Must be removed or disabled before production release.

---

# AI Rules

When generating VBA code, AI must:

- avoid redundant comments
- only document non-obvious logic
- prefer naming over explanation
- always document business rules and constraints
- document workarounds explicitly
- never generate decorative comment blocks
- never leave outdated or commented-out code
- ensure comments remain accurate after code changes
- avoid repeating what the code already expresses
- use concise and professional English

---

# Example

Poor:

```vb
' loop through rows
For i = 1 To 10

    ' set value
    Cells(i, 1).Value = i

Next i
```

Good:

```vb
' Populate test dataset for validation
For i = 1 To 10

    Cells(i, 1).Value = i

Next i
```

---

# Golden Rules

1. Comments explain why, not what.
2. Code should be self-explanatory.
3. Document business rules clearly.
4. Avoid redundant comments.
5. Keep comments accurate and updated.
6. Remove dead or commented-out code.
7. Prefer naming over explanation.
8. Document external constraints.
9. Use TODO/FIXME intentionally.
10. Over-commenting is a design smell.