# 06 - Procedures and Functions

## Objective

Well-designed procedures and functions are the foundation of maintainable VBA code.

Every procedure should have:

- a single responsibility
- a clear purpose
- explicit inputs
- predictable outputs
- minimal side effects

Small, focused procedures are easier to understand, test, reuse and debug.

---

# Single Responsibility Principle

Every procedure should perform **one logical task**.

Good:

```vb
LoadConfiguration
ReadWorksheetData
BuildZplLabel
SavePdf
PrintLabels
```

Avoid:

```vb
ImportAndPrintAndArchiveAndNotify
```

If a procedure name contains multiple verbs, it probably does too much.

---

# Procedure Size

Keep procedures short.

Recommended:

- 10–30 lines: ideal
- 30–60 lines: acceptable
- > 60 lines: review
- > 100 lines: refactor

Long procedures usually contain several responsibilities.

---

# Use Subs for Actions

A `Sub` performs an action.

Examples:

```vb
ExportWorkbook
PrintLabel
RefreshData
InitializeApplication
```

A `Sub` should not return information through global variables.

---

# Use Functions for Results

A `Function` returns a value.

Examples:

```vb
FindLastRow
BuildZplLabel
CalculateWeight
GetPrinterName
```

Functions should avoid modifying external state whenever possible.

---

# Avoid Side Effects

Prefer:

```vb
formattedText = FormatArticle(article)
```

instead of:

```vb
Call FormatArticle(article)
```

where the procedure secretly modifies global variables.

Functions should behave predictably.

---

# Explicit Parameters

Always pass required data explicitly.

Good:

```vb
BuildZplLabel(ws, rowIndex)
```

Avoid:

```vb
BuildZplLabel()
```

that internally reads:

```vb
ActiveSheet
Selection
ActiveCell
```

---

# Parameter Order

Use a consistent order.

Recommended:

1. Main object
2. Context
3. Options
4. Optional parameters

Example:

```vb
ExportWorksheet( _
    ws As Worksheet, _
    outputFolder As String, _
    overwrite As Boolean)
```

---

# Pass the Smallest Required Object

Prefer:

```vb
ws As Worksheet
```

instead of:

```vb
wb As Workbook
```

if only one worksheet is needed.

Avoid passing entire objects when only one value is required.

---

# Avoid Global State

Prefer:

```vb
ProcessOrders(wsOrders)
```

instead of:

```vb
ProcessOrders()
```

that internally depends on:

```vb
gWorkbook
gWorksheet
gConfiguration
```

Explicit dependencies improve readability.

---

# Optional Parameters

Use optional parameters only when they represent true defaults.

Good:

```vb
ExportPdf( _
    fileName As String, _
    Optional openAfterExport As Boolean = False)
```

Avoid procedures with many optional parameters.

---

# Return Values

Functions should return meaningful values.

Good:

```vb
Dim success As Boolean

success = ExportWorkbook(...)
```

Avoid relying solely on:

```vb
ByRef success
```

unless multiple outputs are required.

---

# Multiple Outputs

If several values must be returned:

- use a custom Type
- use a dedicated Class
- or use carefully documented `ByRef` parameters

Avoid returning many unrelated values.

---

# ByVal vs ByRef

Default to:

```vb
ByVal
```

Use:

```vb
ByRef
```

only when the procedure intentionally modifies the caller's variable.

Always make this choice explicit.

---

# Pure Functions

Whenever possible, functions should be pure.

A pure function:

- depends only on its parameters
- always produces the same result
- has no side effects

Example:

```vb
Public Function NormalizeArticle(ByVal article As String) As String
```

Ideal helper functions are pure.

---

# Public Procedures

Public procedures form the project's API.

They should:

- have descriptive names
- validate inputs
- document assumptions
- remain stable over time

Breaking changes should be minimized.

---

# Private Procedures

Private procedures support a single module.

They may assume some internal knowledge but should still follow all coding standards.

---

# Decompose Complex Logic

Instead of:

```vb
ExportLabels()
```

containing 300 lines,

prefer:

```vb
ExportLabels()

    ValidateConfiguration

    LoadOrders

    BuildLabels

    SaveFiles

    PrintLabels

End Sub
```

Each helper performs one task.

---

# Early Exit

Return early for invalid conditions.

Example:

```vb
If ws Is Nothing Then Exit Sub

If lastRow < 2 Then Exit Sub
```

Avoid deeply nested code.

---

# Nesting

Avoid excessive nesting.

Prefer:

```vb
If Not IsValid Then Exit Sub

If Not HasConfiguration Then Exit Sub

GenerateLabels
```

instead of:

```vb
If IsValid Then

    If HasConfiguration Then

        ...

    End If

End If
```

---

# One Level of Abstraction

A procedure should operate at a single level of abstraction.

Avoid mixing:

```vb
LoadOrders()

...

Cells(row, 5).Font.Bold = True
```

The second statement belongs in another helper.

---

# Naming

Procedure names should clearly express intent.

Examples:

```vb
ReadConfiguration
LoadArticles
GenerateLabels
ExportPdf
PrintReport
```

Avoid:

```vb
Run
Execute
Process
DoStuff
```

---

# Comments

Good procedures require very few comments.

If many comments are needed, consider splitting the procedure.

---

# Error Handling

Every non-trivial public procedure should follow the project's error handling policy.

Unexpected errors should be logged using:

```vb
Utils_Log.Error
```

Procedures should restore Excel's state before exiting.

---

# Performance

Do not optimize prematurely.

Prefer readable code.

Optimize only after measuring.

---

# Testing

Small procedures are easier to test.

A procedure performing one task is easier to validate than one performing ten.

---

# AI Rules

When generating VBA code, AI should:

- keep procedures focused on one responsibility
- prefer many small procedures over one large procedure
- pass dependencies explicitly
- avoid global variables
- default parameters to `ByVal`
- use `ByRef` only intentionally
- avoid hidden side effects
- prefer pure helper functions
- use descriptive procedure names
- decompose complex workflows into reusable helpers

---

# Recommended Workflow Example

```text
ExportLabels
├── ValidateConfiguration
├── LoadOrders
├── BuildLabelData
├── GenerateZpl
├── SaveFiles
├── PrintLabels
└── WriteExecutionLog
```

Each procedure has one clear responsibility.

---

# Golden Rules

1. One procedure, one responsibility.
2. Keep procedures short.
3. Use `Sub` for actions.
4. Use `Function` for returned values.
5. Pass dependencies explicitly.
6. Prefer `ByVal` by default.
7. Avoid hidden side effects.
8. Prefer pure helper functions.
9. Break complex workflows into small reusable procedures.
10. Code should read like a sequence of business actions.