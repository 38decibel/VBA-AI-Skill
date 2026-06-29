# 23 - Refactoring

## Objective

Refactoring is the process of improving existing VBA code **without changing its external behavior**.

The goal is to:

- improve readability
- reduce complexity
- increase performance
- enforce architecture rules
- remove duplication
- prepare code for scalability

In VBA, refactoring is not optional — it is a continuous requirement.

---

# Golden Rule

> If code works but is hard to read, slow, or duplicated — it must be refactored.

---

# 1. Refactoring Principles

Refactoring must always preserve:

- business behavior
- outputs
- Excel results
- external integrations

But improve:

- structure
- clarity
- performance
- maintainability

---

# 2. When to Refactor

Refactor when you see:

- duplicated logic
- long procedures (>50 lines)
- deep nesting
- repeated Excel calls
- mixed responsibilities
- unclear naming
- performance issues
- unclear dependencies

---

# 3. Code Smell Indicators

## 3.1 Long procedures

Bad:

```vb
Sub ProcessData()
    ' 200 lines of logic
End Sub
```

---

## 3.2 Duplicate logic

Same code repeated in multiple modules.

---

## 3.3 Mixed responsibilities

One procedure doing:

- Excel read
- business logic
- export
- logging

---

## 3.4 Deep nesting

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

## 3.5 Excessive Excel access

```vb
ws.Cells(i, 1).Value
```

inside loops.

---

# 4. Refactoring Strategy

## Step 1: Extract logic

Move logic into:

- functions
- classes
- service modules

---

## Step 2: Simplify flow

Replace nested conditions with guard clauses.

---

## Step 3: Remove duplication

Centralize repeated logic.

---

## Step 4: Optimize performance

Replace Excel loops with arrays.

---

## Step 5: Rename for clarity

Rename variables and procedures to business meaning.

---

# 5. Extract Method Pattern

## Before

```vb
Sub ProcessOrder()
    ' validation
    ' transformation
    ' export
End Sub
```

---

## After

```vb
Sub ProcessOrder()
    ValidateOrder
    TransformOrder
    ExportOrder
End Sub
```

---

# 6. Extract Class Pattern

If a module grows too large:

Move logic into a class:

```
OrderProcessor → Order class + services
```

---

# 7. Replace Magic Numbers

## Bad

```vb
If qty > 10 Then
```

---

## Good

```vb
If qty > MIN_ORDER_QTY Then
```

---

# 8. Simplify Conditions

## Bad

```vb
If x = True Then
```

---

## Good

```vb
If x Then
```

---

# 9. Remove Duplication

## Bad

Same validation repeated everywhere.

---

## Good

Centralized validation function:

```vb
If Not Validator.IsValidOrder(orderId) Then Exit Sub
```

---

# 10. Improve Naming

## Bad

```vb
Sub DoStuff()
```

## Good

```vb
Sub GenerateZplLabel()
```

---

# 11. Reduce Excel Calls

## Bad

```vb
ws.Cells(i, 1).Value
ws.Cells(i, 2).Value
```

---

## Good

```vb
data = ws.Range("A1:B1000").Value2
```

---

# 12. Replace Long Parameter Lists

## Bad

```vb
Sub CreateOrder(a, b, c, d, e, f)
```

---

## Good

Use class:

```vb
order.Init data
```

---

# 13. Replace Global Variables

Global state must be eliminated or encapsulated.

---

# 14. Refactor for Performance

Key optimizations:

- arrays instead of ranges
- dictionary for lookups
- batch processing
- reduced COM calls

---

# 15. Refactor for Readability

Code must be:

- self-explanatory
- minimal comments needed
- structured in blocks

---

# 16. Refactor for Testability

Break logic into:

- pure functions
- isolated classes
- deterministic outputs

---

# 17. Safe Refactoring Process

## Step 1

Ensure code is working.

## Step 2

Add logging if missing.

## Step 3

Refactor in small steps.

## Step 4

Test after each change.

---

# 18. Refactoring vs Rewrite

## Refactor when:

- logic is correct
- structure is bad

## Rewrite when:

- logic is unclear
- architecture is broken
- performance is catastrophic

---

# 19. Anti-Patterns During Refactoring

- changing business logic accidentally
- refactoring too many things at once
- removing validation logic
- skipping testing

---

# 20. AI Rules

When generating refactored VBA code, AI must:

- preserve behavior exactly
- reduce duplication
- simplify structure
- extract methods logically
- introduce classes when needed
- improve naming clarity
- eliminate Excel cell-by-cell operations
- replace nested conditions with guard clauses
- ensure performance improvements
- maintain logging consistency
- avoid overengineering

---

# Example

## Before

```vb
For i = 1 To lastRow
    If ws.Cells(i, 1).Value <> "" Then
        ws.Cells(i, 2).Value = ws.Cells(i, 1).Value * 2
    End If
Next i
```

---

## After

```vb
data = ws.Range("A1:B" & lastRow).Value2

For i = 1 To UBound(data, 1)
    If data(i, 1) <> "" Then
        data(i, 2) = data(i, 1) * 2
    End If
Next i

ws.Range("A1:B" & lastRow).Value2 = data
```

---

# Golden Rules

1. Refactor continuously.
2. Never change behavior accidentally.
3. Remove duplication aggressively.
4. Simplify structure.
5. Improve naming.
6. Reduce Excel access.
7. Extract methods early.
8. Introduce classes when needed.
9. Keep functions small.
10. Always prioritize readability + performance.