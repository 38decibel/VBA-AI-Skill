# 26 - AI Generation Rules (VBA Code Generation Standard)

## Objective

This chapter defines the **rules an AI must follow when generating VBA code** within this framework.

It ensures:

- architectural consistency
- performance safety
- maintainability
- strict layering
- production-ready output by default

---

# Golden Rule

> AI must generate production-grade VBA code by default, not “example code”.

---

# 1. Mandatory Architecture Compliance

AI MUST respect the full layering model:

```
UserForm → Controller Module → Class Modules → Excel Object Model
                                   ↓
                                Utils_Log
```

---

## Forbidden violations

- business logic in UserForms
- Excel manipulation in events
- UI logic in classes
- COM usage in business layer
- mixed responsibilities in modules

---

# 2. Default Design Philosophy

AI must assume:

- Excel is unstable
- input data is untrusted
- performance matters
- code will be maintained long-term

---

# 3. Performance-First Rule

AI MUST:

- use `.Value2` for all Excel reads/writes
- use arrays for bulk processing
- avoid cell-by-cell operations
- minimize Excel object calls
- avoid loops over ListRows when arrays are possible

---

# 4. Defensive Programming is Mandatory

AI MUST:

- validate all inputs
- check `Nothing` before use
- validate arrays with bounds checks
- use `.Exists` for dictionaries
- handle empty ranges safely
- implement full error handling

---

# 5. Error Handling Standard

Every procedure MUST include:

```vb
On Error GoTo ErrorHandler

' logic

Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "Module.Procedure"
```

No exceptions.

---

# 6. Logging Rules

AI MUST:

- log key operations (Info/Debug)
- log all errors with context
- avoid MsgBox for debugging
- use structured format:

```vb
Utils_Log.Debug "Module", "Message"
```

---

# 7. Naming Conventions (Strict)

AI MUST:

- use meaningful business names
- avoid vague names (`DoStuff`, `ProcessData`)
- use PascalCase for procedures
- prefix private fields with `p`

---

# 8. Excel Object Model Rules

AI MUST:

- avoid `ActiveSheet`, `ActiveWorkbook`
- avoid `Select` and `Activate`
- use explicit worksheet references
- use ListObjects for structured data
- treat Excel as I/O layer only

---

# 9. COM Automation Rules

AI MUST:

- use late binding unless explicitly required
- release all COM objects (`Set obj = Nothing`)
- avoid COM inside loops
- encapsulate COM in service modules
- never mix COM with business logic

---

# 10. File System Rules

AI MUST:

- never hardcode file paths
- validate file existence before use
- use `FileSystemObject`
- sanitize file names
- close all file handles

---

# 11. Event Handling Rules

AI MUST:

- keep events minimal
- delegate to controllers
- avoid loops in event handlers
- manage `EnableEvents` safely
- prevent reentrancy issues

---

# 12. UserForm Rules

AI MUST:

- keep UI logic only in forms
- delegate processing to controllers
- avoid Excel access in forms
- avoid business logic in UI
- keep event handlers small

---

# 13. Class Design Rules

AI MUST:

- use classes for business entities
- enforce single responsibility
- keep fields private
- use Init pattern instead of constructors
- avoid God classes

---

# 14. Anti-Pattern Avoidance (STRICT)

AI MUST NEVER generate:

- ActiveSheet usage
- Select / Activate
- cell-by-cell loops for large data
- silent error handling (`Resume Next`)
- global variables without justification
- hardcoded paths
- magic numbers
- COM leaks
- business logic in UI/events

---

# 15. Refactoring Mindset

AI MUST:

- prefer modular design
- extract repeated logic into functions/classes
- simplify nested conditions (guard clauses)
- reduce duplication
- optimize for readability and performance

---

# 16. Data Handling Rules

AI MUST:

- use arrays for Excel data
- use Dictionary for lookups
- avoid Collection for complex logic
- assume Excel data is Variant
- validate bounds before access

---

# 17. Performance Rules

AI MUST:

- disable ScreenUpdating for heavy tasks
- disable Events during batch updates
- use Calculation manual when needed
- batch read/write operations
- avoid repeated Excel calls in loops

---

# 18. Output Quality Rules

AI MUST generate:

- clean, structured code
- production-ready logic
- minimal but sufficient comments
- consistent naming
- maintainable architecture

---

# 19. Context Awareness Rule

AI MUST:

- respect existing module architecture
- not introduce unnecessary layers
- not duplicate existing utilities
- reuse existing patterns (e.g., Utils_Log)

---

# 20. Safety Rule

AI MUST:

- assume all external systems are unreliable
- validate all inputs
- handle all failures gracefully
- never crash silently

---

# 21. Minimal Viable Complexity Rule

AI MUST:

> choose the simplest solution that is still correct, safe, and performant.

Avoid overengineering.

---

# 22. Example of Correct AI Output

```vb
Public Sub ProcessOrders()

    On Error GoTo ErrorHandler

    Dim data As Variant
    data = lo.DataBodyRange.Value2

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If data(i, 1) <> "" Then
            data(i, 2) = data(i, 1) * 2
        End If
    Next i

    lo.DataBodyRange.Value2 = data

    Utils_Log.Info "ProcessOrders", "Completed"

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "ProcessOrders"
End Sub
```

---

# 23. AI Output Checklist

Before finalizing code, AI must ensure:

- [ ] no ActiveSheet / Select usage
- [ ] arrays used for bulk data
- [ ] error handling present
- [ ] logging included
- [ ] no business logic in UI/events
- [ ] no COM leaks
- [ ] performance optimized
- [ ] naming is meaningful
- [ ] architecture respected
- [ ] defensive programming applied

---

# Golden Rules

1. Generate production-ready code by default.
2. Respect strict architecture layers.
3. Optimize for performance and readability.
4. Always include error handling.
5. Always include logging.
6. Never use Excel as a processor.
7. Always use arrays for data.
8. Avoid all anti-patterns.
9. Ensure safe external system usage.
10. Prefer simplicity over complexity.