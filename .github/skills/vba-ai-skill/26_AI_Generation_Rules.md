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

> AI must generate production-grade VBA code by default, not â€œexample codeâ€.

---

# 1. Mandatory Architecture Compliance

AI MUST respect the full layering model:

```
UserForm â†’ Controller Module â†’ Class Modules â†’ Excel Object Model
                                   â†“
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
- check `Nothing` before use: raise through `Utils_Guard` for contract violations, never exit silently
- validate arrays with bounds checks
- use `.Exists` for dictionaries
- handle empty ranges safely
- apply the error-handling policy of chapter 05 according to the procedure role

---

# 5. Error Handling Standard

The handler policy depends on the **role** of the procedure (chapter 05):

| Role | Handler |
|------|---------|
| Entry point (macro, event, button, `Workbook_Open`) | Always: log once with `Utils_Log.Error`, never let the error escape |
| Owns a resource or application state | Clean up in `CleanExit`, then re-raise |
| Plain helper, `Try` function | No handler; guard clauses raise through `Utils_Guard` |

Standard structure for an entry point:

```vb
On Error GoTo CleanFail

' logic

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Module.Procedure"
    Resume CleanExit
```

Labels are always `CleanExit` and `CleanFail`. Never use `Call`, never `On Error GoTo 0` inside a procedure
with a handler, never release resources outside `CleanExit`.

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
- use English identifiers
- keep a consistent naming style in each module (`PascalCase` or `lowerCamelCase`)
- avoid type-based Hungarian notation; `m_` / `g_` scope prefixes, `cls` for classes, `I` for interfaces and `out` for output parameters are the project conventions
- reserve underscores in procedure names for event handlers and `Interface_Member`
- declare each variable where it is first used, one identifier = one purpose

---

# 8. Excel Object Model Rules

AI MUST:

- avoid `ActiveSheet`, `ActiveWorkbook`
- avoid `Select` and `Activate`
- use `Worksheets` rather than `Sheets`; read and write `.Value2` explicitly
- use explicit worksheet references
- use ListObjects for structured data
- treat Excel as I/O layer only

---

# 9. COM Automation Rules

AI MUST:

- use explicit late binding for delivered code (`As Object` + `CreateObject`); early binding is for development
- release all COM objects in `CleanExit`
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
- create forms with `New`, never show the default instance; handle `QueryClose` and hide the form
- keep form data in a model class with `IsValid` / `IsCancelled`
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
- use an `Init` method, or a `Create` factory returning an interface (chapter 28), instead of constructors
- name classes `clsXxx` and interfaces `IXxx`; inject dependencies instead of creating them
- write unit-testable code and propose tests for business rules (chapter 29)
- avoid God classes

---

# 14. Anti-Pattern Avoidance (STRICT)

AI MUST NEVER generate:

- ActiveSheet usage
- Select / Activate
- cell-by-cell loops for large data
- silent error handling (`Resume Next` without immediate restore/validation/logging)
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

    On Error GoTo CleanFail

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

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "ProcessOrders"
    Resume CleanExit

End Sub
```

---

# 23. AI Output Checklist

Before finalizing code, AI must ensure:

- [ ] no ActiveSheet / Select usage
- [ ] arrays used for bulk data
- [ ] error handling matches the procedure role (CleanExit / CleanFail)
- [ ] no `Call`, no obsolete constructs, every parameter explicit `ByVal` / `ByRef`
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
