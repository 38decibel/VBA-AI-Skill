# 25 - Anti-Patterns (VBA Code Smells & Forbidden Practices)

## Objective

This chapter defines **strictly forbidden patterns** in VBA development.

Anti-patterns are not just â€œbad styleâ€ â€” they are:

- performance risks
- maintenance blockers
- sources of hidden bugs
- architectural violations
- scalability killers

---

# Golden Rule

> If you recognize an anti-pattern in production code, it must be refactored immediately.

---

# 1. Excel Object Abuse

## 1.1 ActiveSheet / ActiveWorkbook

### âŒ Forbidden

```vb
ActiveSheet.Range("A1").Value = 1
```

### Why it's bad

- non-deterministic
- breaks automation
- depends on UI state

### âœ” Correct

```vb
ws.Range("A1").Value2 = 1
```

---

## 1.2 Select / Activate

### âŒ Forbidden

```vb
Range("A1").Select
Selection.Value = 10
```

### Why it's bad

- slow
- unstable
- UI-dependent

---

# 2. Cell-by-Cell Processing

## âŒ Anti-pattern

```vb
For i = 1 To lastRow
    ws.Cells(i, 1).Value = ws.Cells(i, 2).Value
Next i
```

---

## Why it's critical

- extremely slow
- triggers Excel overhead per call
- impossible to scale

---

## âœ” Correct

```vb
data = ws.Range("A1:B" & lastRow).Value2
```

---

# 3. ReDim in Loops

## âŒ Anti-pattern

```vb
For i = 1 To 10000
    ReDim Preserve arr(i)
Next i
```

---

## Why it's bad

- O(nÂ²) complexity
- memory fragmentation
- extremely slow

---

## âœ” Correct

Pre-size arrays or use Dictionary.

---

# 4. Silent Error Handling

## âŒ Anti-pattern

```vb
On Error Resume Next
```

without control.

---

## Why it's dangerous

- hides bugs
- corrupts data silently
- breaks debugging

---

## âœ” Correct

```vb
On Error GoTo ErrorHandler
```

---

# 5. Business Logic in UI Layer

## âŒ Anti-pattern

Inside UserForms:

```vb
If qty > 10 Then discount = 0.2
```

---

## Why it's bad

- violates architecture
- untestable
- duplicated logic risk

---

## âœ” Correct

```vb
discount = order.CalculateDiscount(qty)
```

---

# 6. God Modules

## âŒ Anti-pattern

Modules that:

- contain everything
- mix responsibilities
- grow indefinitely

Examples:

- `Module1`
- `UtilsEverything`
- `Manager`

---

## Why it's bad

- unmaintainable
- impossible to refactor safely

---

# 7. God Classes

## âŒ Anti-pattern

Classes that:

- handle multiple domains
- contain Excel + business + logging + UI logic

---

## Example bad class

- `ExcelProcessor`
- `DataManager`
- `AppController`

---

## âœ” Correct

Split into:

- `Order`
- `OrderValidator`
- `OrderExporter`

---

# 8. Global State Abuse

## âŒ Anti-pattern

```vb
Public currentOrder As String
```

---

## Why it's bad

- unpredictable behavior
- hidden dependencies
- concurrency issues

---

## âœ” Correct

Encapsulate in classes.

---

# 9. Hardcoded Paths

## âŒ Anti-pattern

```vb
"C:\Users\John\Desktop\file.xlsx"
```

---

## Why it's bad

- non-portable
- breaks in production
- environment dependent

---

## âœ” Correct

```vb
basePath & "\file.xlsx"
```

---

# 10. Magic Numbers

## âŒ Anti-pattern

```vb
If qty > 10 Then
```

---

## âœ” Correct

```vb
Const MAX_QTY As Long = 10
```

---

# 11. Repeated Excel Calls

## âŒ Anti-pattern

```vb
ws.Range("A1").Value
ws.Range("A1").Value
ws.Range("A1").Value
```

---

## Why it's bad

- redundant COM overhead
- performance degradation

---

## âœ” Correct

Cache values in variables or arrays.

---

# 12. Deep Nesting

## âŒ Anti-pattern

```vb
If a Then
    If b Then
        If c Then
            If d Then
            End If
        End If
    End If
End If
```

---

## âœ” Correct

Use guard clauses:

```vb
If Not a Then Exit Sub
If Not b Then Exit Sub
```

---

# 13. Overuse of MsgBox

## âŒ Anti-pattern

- debugging via MsgBox
- system flow interruptions

---

## Why it's bad

- blocks execution
- not scalable
- unprofessional logging

---

## âœ” Correct

```vb
Utils_Log.Debug "Message"
```

---

# 14. No Error Context

## âŒ Anti-pattern

```vb
Utils_Log.Error Err
```

---

## âœ” Correct

```vb
Utils_Log.Error Err, "Module.Procedure"
```

---

# 15. Unreleased COM Objects

## âŒ Anti-pattern

```vb
Set olApp = CreateObject("Outlook.Application")
' no cleanup
```

---

## Why it's critical

- memory leaks
- zombie processes

---

## âœ” Correct

```vb
Set olApp = Nothing
```

---

# 16. Overuse of Variant

## âŒ Anti-pattern

```vb
Dim x As Variant
```

without reason.

---

## Why it's bad

- weak typing
- hidden errors
- performance cost

---

## âœ” Correct

Use explicit types whenever possible.

---

# 17. Uncontrolled Loops

## âŒ Anti-pattern

```vb
Do While True
```

without exit condition.

---

# 18. Using Excel as a Processor

## âŒ Anti-pattern

- formulas instead of logic layer
- repeated worksheet recalculations
- heavy dependency on Excel engine

---

## âœ” Correct

Process in memory using arrays.

---

# 19. No Logging Strategy

## âŒ Anti-pattern

- no traceability
- silent execution

---

## âœ” Correct

Use structured logging:

```vb
Utils_Log.Info "Module", "Action started"
```

---

# 20. Mixing Layers

## âŒ Anti-pattern

- UI + business + data in same procedure
- COM + Excel + logic combined

---

## âœ” Correct

Follow strict layering:

```
UI â†’ Controller â†’ Class â†’ Excel
```

---

# AI Rules

When generating VBA code, AI must:

- avoid ALL listed anti-patterns
- enforce architecture separation
- never use ActiveSheet/Select
- always use arrays for bulk operations
- always include error handling
- avoid global state
- avoid ReDim in loops
- use explicit typing
- enforce logging context
- prevent COM leaks
- ensure performance-safe patterns

---

# Final Warning

> Any occurrence of these anti-patterns in production code is considered a critical defect.

---

# Golden Rules

1. No ActiveSheet / Select usage.
2. No cell-by-cell processing for large data.
3. No silent error handling.
4. No global state abuse.
5. No business logic in UI.
6. No God modules or classes.
7. No hardcoded paths.
8. No magic numbers.
9. No COM leaks.
10. No mixed-layer logic.
