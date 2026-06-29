# 25 - Anti-Patterns (VBA Code Smells & Forbidden Practices)

## Objective

This chapter defines **strictly forbidden patterns** in VBA development.

Anti-patterns are not just “bad style” — they are:

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

### ❌ Forbidden

```vb
ActiveSheet.Range("A1").Value = 1
```

### Why it's bad

- non-deterministic
- breaks automation
- depends on UI state

### ✔ Correct

```vb
ws.Range("A1").Value2 = 1
```

---

## 1.2 Select / Activate

### ❌ Forbidden

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

## ❌ Anti-pattern

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

## ✔ Correct

```vb
data = ws.Range("A1:B" & lastRow).Value2
```

---

# 3. ReDim in Loops

## ❌ Anti-pattern

```vb
For i = 1 To 10000
    ReDim Preserve arr(i)
Next i
```

---

## Why it's bad

- O(n²) complexity
- memory fragmentation
- extremely slow

---

## ✔ Correct

Pre-size arrays or use Dictionary.

---

# 4. Silent Error Handling

## ❌ Anti-pattern

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

## ✔ Correct

```vb
On Error GoTo ErrorHandler
```

---

# 5. Business Logic in UI Layer

## ❌ Anti-pattern

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

## ✔ Correct

```vb
discount = order.CalculateDiscount(qty)
```

---

# 6. God Modules

## ❌ Anti-pattern

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

## ❌ Anti-pattern

Classes that:

- handle multiple domains
- contain Excel + business + logging + UI logic

---

## Example bad class

- `ExcelProcessor`
- `DataManager`
- `AppController`

---

## ✔ Correct

Split into:

- `Order`
- `OrderValidator`
- `OrderExporter`

---

# 8. Global State Abuse

## ❌ Anti-pattern

```vb
Public currentOrder As String
```

---

## Why it's bad

- unpredictable behavior
- hidden dependencies
- concurrency issues

---

## ✔ Correct

Encapsulate in classes.

---

# 9. Hardcoded Paths

## ❌ Anti-pattern

```vb
"C:\Users\John\Desktop\file.xlsx"
```

---

## Why it's bad

- non-portable
- breaks in production
- environment dependent

---

## ✔ Correct

```vb
basePath & "\file.xlsx"
```

---

# 10. Magic Numbers

## ❌ Anti-pattern

```vb
If qty > 10 Then
```

---

## ✔ Correct

```vb
Const MAX_QTY As Long = 10
```

---

# 11. Repeated Excel Calls

## ❌ Anti-pattern

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

## ✔ Correct

Cache values in variables or arrays.

---

# 12. Deep Nesting

## ❌ Anti-pattern

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

## ✔ Correct

Use guard clauses:

```vb
If Not a Then Exit Sub
If Not b Then Exit Sub
```

---

# 13. Overuse of MsgBox

## ❌ Anti-pattern

- debugging via MsgBox
- system flow interruptions

---

## Why it's bad

- blocks execution
- not scalable
- unprofessional logging

---

## ✔ Correct

```vb
Utils_Log.Debug "Message"
```

---

# 14. No Error Context

## ❌ Anti-pattern

```vb
Utils_Log.Error Err
```

---

## ✔ Correct

```vb
Utils_Log.Error Err, "Module.Procedure"
```

---

# 15. Unreleased COM Objects

## ❌ Anti-pattern

```vb
Set olApp = CreateObject("Outlook.Application")
' no cleanup
```

---

## Why it's critical

- memory leaks
- zombie processes

---

## ✔ Correct

```vb
Set olApp = Nothing
```

---

# 16. Overuse of Variant

## ❌ Anti-pattern

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

## ✔ Correct

Use explicit types whenever possible.

---

# 17. Uncontrolled Loops

## ❌ Anti-pattern

```vb
Do While True
```

without exit condition.

---

# 18. Using Excel as a Processor

## ❌ Anti-pattern

- formulas instead of logic layer
- repeated worksheet recalculations
- heavy dependency on Excel engine

---

## ✔ Correct

Process in memory using arrays.

---

# 19. No Logging Strategy

## ❌ Anti-pattern

- no traceability
- silent execution

---

## ✔ Correct

Use structured logging:

```vb
Utils_Log.Info "Module", "Action started"
```

---

# 20. Mixing Layers

## ❌ Anti-pattern

- UI + business + data in same procedure
- COM + Excel + logic combined

---

## ✔ Correct

Follow strict layering:

```
UI → Controller → Class → Excel
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