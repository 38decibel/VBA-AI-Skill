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

# Procedure Size and Complexity

Keep procedures short.

Recommended (single reference thresholds for the whole skill):

- 10-30 lines: ideal
- 30-50 lines: acceptable
- more than 50 lines: review
- more than 100 lines: refactor

Long procedures usually contain several responsibilities.

Line count is a signal, not a goal: never split a procedure arbitrarily to hit a number, and never hurt
readability to save lines. Complement it with two other signals:

- **Nesting depth**: two nested loops over a 2D structure (rows and columns) are reasonable.
  Beyond that, make the nesting implicit with a procedure call.
- **Cyclomatic complexity** (number of independent execution paths): a procedure above ~5 is harder to follow
  than one at 1 or 2. Procedures with deeply nested loops and conditions that reach 20 or more must be split.

Rule of thumb: **extracting the body of a loop into its own parameterized procedure is almost always a good idea**.
The arrow-shaped code flattens, the line count drops, and each procedure has fewer reasons to fail.

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

# Calling Procedures: `Call` and Parentheses

## No `Call` keyword

`Call` is obsolete. Use expressive procedure names instead; the verb in the name is the action.

Good:

```vb
ExportWorksheet ws, outputFolder, overwriteExisting
```

Never:

```vb
Call ExportWorksheet(ws, outputFolder, overwriteExisting)
```

## Parentheses rules

- **Ignoring the return value, or calling a `Sub`**: no parentheses around the argument list.

  ```vb
  MsgBox "Export completed", vbInformation
  Utils_Log.Info "Export", "Completed"
  ```

- **Capturing the return value** (assigning it, or passing it as an argument to another call): parentheses are required.

  ```vb
  answer = MsgBox("Continue?", vbYesNo)
  lastRow = FindLastRow(ws)
  ```

- Never put the whole argument list of a `Sub` in parentheses with several arguments: it does not compile
  (`MsgBox ("Hello", "Title")` is an error).

## Parentheses around a single argument pass a copy

```vb
DoSomething (someVariable)
```

evaluates the expression and passes its result instead of a reference to the variable. This is occasionally
useful to protect a `ByRef` variable, but it is almost always a mistake. Never add parentheses around a single
argument unless a copy is intended and obvious.

---

# Avoid Side Effects

Prefer:

```vb
formattedText = FormatArticle(article)
```

over a procedure that secretly modifies global variables.

Functions should behave predictably.

---

# Explicit Parameters

Always pass required data explicitly.

Good:

```vb
BuildZplLabel ws, rowIndex
```

Avoid a parameterless procedure that internally reads:

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
Public Sub ExportWorksheet( _
    ByVal ws As Worksheet, _
    ByVal outputFolder As String, _
    ByVal overwrite As Boolean)
```

---

# Pass the Smallest Required Object

Prefer:

```vb
ByVal ws As Worksheet
```

instead of:

```vb
ByVal wb As Workbook
```

if only one worksheet is needed.

Avoid passing entire objects when only one value is required.

---

# Avoid Global State

Prefer passing `wsOrders` to `ProcessOrders` as a parameter instead of letting it depend on:

```vb
g_Workbook
g_Worksheet
g_Configuration
```

Explicit dependencies improve readability.

Do not promote a local variable to module level just because two procedures need it: pass it as a parameter.

---

# ByVal, ByRef and Named Arguments

## Every parameter has an explicit modifier

VBA passes arguments `ByRef` by default. Never rely on that default:

```vb
Public Function FindLastRow(ByVal ws As Worksheet) As Long
```

Always write `ByVal` or `ByRef`, even for object parameters.

## Default to `ByVal`

Use `ByRef` only when the procedure intentionally modifies the caller's variable
(for example a `Try` function that assigns an output parameter).

Objects are never really "passed": what travels is a pointer, even with `ByVal` (a copy of the pointer).
Passing an object `ByRef` additionally lets the callee reassign the caller's variable, which increases the
chances of programming errors.

## Arrays and user-defined types

Arrays and `Type` structures cannot be passed `ByVal` (unless wrapped in a `Variant`). They must be `ByRef`.
To pass an array between scopes, a `Variant` parameter is often the simplest option.

## Output parameters: the `out` prefix

VBA has no `Out` keyword. Mark `ByRef` output parameters with an `out` prefix so the intent is obvious
at the call site and in the signature:

```vb
Public Function TryParseQuantity(ByVal text As String, ByRef outQuantity As Long) As Boolean
```

## Named arguments

Use named arguments when:

- a literal whose meaning is not obvious is passed (`True`, `False`, magic numbers)
- an `out` parameter is passed
- several optional parameters are skipped

```vb
ExportPdf fileName, openAfterExport:=True

If Not Utils_Excel.TryGetWorksheet(wb, "Orders", outWorksheet:=wsOrders) Then
```

Keep a name and its `:=` operator and value on the same line when splitting long argument lists.

---

# Optional Parameters

Use optional parameters only when they represent true defaults.

Good:

```vb
Public Sub ExportPdf( _
    ByVal fileName As String, _
    Optional ByVal openAfterExport As Boolean = False)
```

Avoid procedures with many optional parameters.

---

# Return Values

Functions should return meaningful values.

Good:

```vb
Dim success As Boolean
success = ExportWorkbook(wb)
```

Avoid relying solely on `ByRef success` unless multiple outputs are required.

---

# Multiple Outputs and the Try Pattern

If several values must be returned:

- use a custom Type
- use a dedicated Class
- or use carefully documented `ByRef` parameters with the `out` prefix

When an operation can **expectedly** fail, expose it as a `TryXxx` function that returns `Boolean`
and delivers the result through an `out` parameter (see chapter 05). This keeps `On Error Resume Next`
confined to one tiny procedure.

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

Instead of one `ExportLabels` procedure containing 300 lines, prefer:

```vb
Public Sub ExportLabels()

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

Return early for invalid conditions:

- a contract violation raises an error through `Utils_Guard` (chapter 05)
- an expected data situation logs a `Warning` and exits

```vb
Utils_Guard.NotNothing ws, "ws"

If lastRow < 2 Then Exit Sub
```

Avoid deeply nested code.

---

# Nesting

Avoid excessive nesting.

Prefer:

```vb
If Not isValid Then Exit Sub
If Not hasConfiguration Then Exit Sub

GenerateLabels
```

instead of:

```vb
If isValid Then

    If hasConfiguration Then

        ' ...

    End If

End If
```

---

# One Level of Abstraction

A procedure should operate at a single level of abstraction.

Avoid mixing:

```vb
LoadOrders

Cells(row, 5).Font.Bold = True
```

The second statement belongs in another helper.

Cells, ranges and Win32 API calls live at the lowest level; higher levels express intent
("prepare the report header") rather than mechanics ("write a series of values in a series of cells").

---

# Module Layout: One Macro per Module

Keep a macro (a public entry point without parameters) in its own standard module, not in a worksheet code-behind.
Order the module like a story:

1. the public procedure at the top, at the highest level of abstraction
2. private parameterized procedures below, in the order they are called, at decreasing levels of abstraction

The top reads like a summary; the bottom reads like a series of small, specific and rather boring operations.

You know it is well done when a second macro needs the same low-level operations: extract them into public
members of a utility module (`Option Private Module`) and reuse them. Never copy and paste a block into another module.

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

Follow the procedure roles defined in chapter 05:

- entry points always have a handler that logs once with `Utils_Log.Error`
- procedures that own a resource or an application state clean up in `CleanExit` and re-raise
- plain helpers have no handler and rely on guard clauses

---

# Performance

Do not optimize prematurely.

Prefer readable code.

Optimize only after measuring.

---

# Testing

Small procedures are easier to test.

A procedure performing one task is easier to validate than one performing ten.
See chapter 29 for unit testing principles.

---

# AI Rules

When generating VBA code, AI should:

- keep procedures focused on one responsibility
- prefer many small procedures over one large procedure
- extract loop bodies into their own parameterized procedures
- pass dependencies explicitly
- avoid global variables
- never use the `Call` keyword
- use parentheses only when capturing a return value
- write an explicit `ByVal` or `ByRef` on every parameter, defaulting to `ByVal`
- pass arrays and user-defined types `ByRef`
- prefix `ByRef` output parameters with `out` and pass them as named arguments
- expose expected failures as `TryXxx` functions
- avoid hidden side effects
- prefer pure helper functions
- use descriptive procedure names
- put the highest-level procedure first, then private helpers by decreasing abstraction
- decompose complex workflows into reusable helpers

---

# Recommended Workflow Example

```text
ExportLabels
|-- ValidateConfiguration
|-- LoadOrders
|-- BuildLabelData
|-- GenerateZpl
|-- SaveFiles
|-- PrintLabels
`-- WriteExecutionLog
```

Each procedure has one clear responsibility.

---

# Golden Rules

1. One procedure, one responsibility.
2. Keep procedures short and shallow.
3. Use `Sub` for actions.
4. Use `Function` for returned values.
5. Pass dependencies explicitly.
6. Every parameter is explicitly `ByVal` (default) or `ByRef` (intentional).
7. No `Call`; parentheses only to capture a return value.
8. Avoid hidden side effects.
9. Prefer pure helper functions.
10. Code should read like a sequence of business actions.
