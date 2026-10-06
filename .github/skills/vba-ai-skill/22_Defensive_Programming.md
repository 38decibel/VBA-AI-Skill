# 22 - Defensive Programming

## Objective

Defensive programming is about building VBA code that:

- fails safely instead of crashing
- fails early and loudly when the caller breaks a contract
- validates inputs early
- assumes external data is unreliable
- prevents silent corruption
- improves traceability and robustness

In VBA, where type safety and compiler guarantees are limited, defensive programming is **mandatory, not optional**.

---

# Golden Rule

> Never trust inputs: Excel data, user input, files, or external systems are all potentially invalid.

Treat cell values exactly like user input: they can be empty, of the wrong type, or formatted differently than expected.

---

# 1. Validate Early, Fail Fast

## Principle

All validation happens at the **entry of a procedure**, before any work is done.

Two kinds of problems exist (see chapter 05):

- **Programming errors** (the caller violated the contract): the guard clause **raises** a custom error.
- **User / data errors** (expected in real life): validate, log a `Warning`, report clearly, exit cleanly.

A silent `Exit Sub` is only acceptable for the second kind, and only together with a log entry.

---

## Example: programming errors

```vb
Utils_Guard.NotNothing ws, "ws"
Utils_Guard.NotEmpty orderId, "orderId"
Utils_Guard.NotNegative quantity, "quantity"
```

## Example: expected data situations

```vb
If lastRow < 2 Then

    Utils_Log.Warning "ExportOrders", "No data rows to export", "Sheet=" & ws.Name
    Exit Sub

End If
```

---

# 2. Null / Nothing Safety

## Always check object parameters

```vb
Utils_Guard.NotNothing lo, "lo"
Utils_Guard.NotNothing rng, "rng"
Utils_Guard.NotNothing wb, "wb"
```

---

## ListObject special case

An empty table has no data body. This is a data situation, not a programming error.

```vb
If lo.DataBodyRange Is Nothing Then

    Utils_Log.Warning "ProcessTable", "Table has no data rows", "Table=" & lo.Name
    Exit Sub

End If
```

---

# 3. Range Safety Rules

Excel ranges can be empty, invalid, or unexpected.

When a range must have a specific shape, verify it:

- a single cell when the code expects one
- a single `Area` when the code is not area-aware
- the expected number of rows and columns
- a real `Range` when handling `Selection` (never assume `Selection` is a range)

## Safe pattern in an event

```vb
If Target.CountLarge > 1 Then Exit Sub
```

---

# 4. Type Safety (Variant danger control)

VBA is loosely typed: defensive checks are required.

## Example

```vb
If Not IsNumeric(value) Then Exit Sub
If Not IsDate(value) Then Exit Sub
```

`IsNumeric`, `CDbl` and `CDate` depend on the user's locale (decimal separator, date format).
Never assume that numbers typed by users, or text imported from files, use the same separators as your machine.
When the format is fixed (SAP exports, CSV files), parse it explicitly instead of relying on implicit locale conversion.

---

# 5. Error Handling Follows Procedure Roles

Defensive code is not "a handler in every procedure". Follow the roles defined in chapter 05:

- entry points always have a handler
- procedures owning a resource or application state have a handler (clean up, then re-raise)
- plain helpers have no handler and rely on guard clauses

```vb
On Error GoTo CleanFail

' logic here

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "ModuleName.ProcedureName"
    Resume CleanExit
```

---

# 6. Never Ignore Errors

Bad:

```vb
On Error Resume Next
```

Without control.

---

## Allowed usage

Only in controlled blocks, preferably inside a dedicated `Try` function (chapter 05):

```vb
Public Function TryOpenWorkbook(ByVal path As String, ByRef outWorkbook As Workbook) As Boolean

    Set outWorkbook = Nothing

    On Error Resume Next
    Set outWorkbook = Workbooks.Open(path)
    On Error GoTo 0

    TryOpenWorkbook = Not outWorkbook Is Nothing

End Function
```

Inside a procedure that has its own handler, restore it with `On Error GoTo CleanFail`, never `On Error GoTo 0`.

---

# 7. Defensive COM Handling

COM objects must always be protected:

```vb
If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
End If
```

COM creation and calls can fail for many reasons outside your control: handle and clean up (chapter 20).

---

# 8. File System Safety

Always validate before access:

```vb
If Not fso.FileExists(path) Then

    Utils_Log.Warning "ReadFile", "File not found", "Path=" & path
    Exit Sub

End If
```

Validating existence does not remove the need to handle unexpected I/O failures
(locked file, network drive unmounted, missing permissions).

---

# 9. Array Safety Rules

Excel arrays must be validated before use:

```vb
If IsEmpty(data) Then Exit Sub
If Not IsArray(data) Then Exit Sub
```

Remember that `Range.Value2` on a single cell returns a scalar, not a 2D array.

---

## Bounds safety

```vb
If UBound(data, 1) < LBound(data, 1) Then Exit Sub
```

---

# 10. Dictionary Safety Rules

Always check existence:

```vb
If dict.Exists(key) Then
    value = dict(key)
End If
```

Never:

```vb
value = dict(key)
```

without validation (reading a missing key silently adds it).

---

# 11. User Input Safety

All UserForm inputs must be validated, and the code must assume the user will do everything possible to break it:

- empty required fields
- values outside the valid range or in the wrong format
- decimal separators and date formats that differ from yours
- a cancelled dialog or a form closed with the [X] button

## Required fields

```vb
If Len(txtOrderId.Value) = 0 Then Exit Sub
```

## Numeric fields

```vb
If Not IsNumeric(txtQty.Value) Then Exit Sub
```

Prefer an `IsValid` property on the form model (chapter 16) that drives the enabled state of the Accept button.

User errors are reported with a clear and specific message ("quantity must be a positive number"),
never with a raw technical error.

---

# 12. Excel State Protection

Always protect Excel environment state:

```vb
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
```

And restore in `CleanExit`:

```vb
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
```

---

# 13. Safe Exit Pattern

Use centralized cleanup (chapter 05):

```vb
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
```

---

# 14. Guard Clauses (MANDATORY)

Avoid deep nesting.

## Bad

```vb
If condition1 Then
    If condition2 Then
        If condition3 Then
            ' logic
        End If
    End If
End If
```

---

## Good

Contract checks raise; data checks exit with a log entry:

```vb
Utils_Guard.NotNothing ws, "ws"

If Not hasRows Then Exit Sub
If Not hasConfiguration Then Exit Sub
```

---

# 15. Defensive Logging

Every handled failure must be traceable:

```vb
Utils_Log.Error Err, "Order.Process"
```

Include context:

```vb
Utils_Log.Error Err, "Order.Process", "OrderId=" & orderId
```

Log once, where the error is finally handled (chapter 05).

---

# 16. Silent Failure is Forbidden

Never swallow errors:

Bad:

```vb
On Error Resume Next
```

without handling.

Bad:

```vb
If ws Is Nothing Then Exit Sub
```

when `Nothing` can only be the result of a programming mistake: raise instead.

---

# 17. Input Sanitization Rules

Always clean inputs:

```vb
value = Trim(value)
value = Replace(value, vbNullChar, "")
```

---

# 18. File Input Validation

Before processing external data:

- check file exists
- check extension
- validate structure

---

# 19. Defensive Looping

Always protect loops:

```vb
If UBound(data, 1) < LBound(data, 1) Then Exit Sub
```

---

# 20. Defensive Design Pattern

Each procedure follows:

```
Validate -> Process -> Validate -> Output -> Log
```

---

# 21. Common Failure Points

## 21.1 Empty ranges

```vb
lo.DataBodyRange
```

may be Nothing.

---

## 21.2 Missing dictionary keys

Always check `.Exists`.

---

## 21.3 Unexpected types from Excel

Everything from Excel is Variant: validate.

---

## 21.4 External systems

COM, files, APIs: always fail-prone.

---

# 22. Defensive vs Overengineering

## Good defensive programming

- validation
- error handling where it does real work
- safe defaults

## Bad overengineering

- excessive checks everywhere
- redundant validation
- handlers that only log and re-raise
- performance degradation

---

# 23. AI Rules

When generating VBA code, AI must:

- always validate inputs
- raise a custom error for contract violations instead of a silent `Exit Sub`
- exit with a `Warning` for expected data situations
- always check objects for Nothing
- put handlers in entry points and in procedures that own resources (not in every procedure)
- never assume Excel data is valid
- use guard clauses instead of nested logic
- ensure COM safety checks
- validate arrays before access
- validate dictionary keys
- protect Excel application state
- ensure safe exits with cleanup

---

# Example

## Bad

```vb
value = dict(key)
```

---

## Good

```vb
If dict.Exists(key) Then
    value = dict(key)
Else
    Utils_Log.Debug "ProcessOrders", "Missing key: " & key
End If
```

---

# Golden Rules

1. Never trust inputs.
2. Validate early and often.
3. Contract violations raise; expected data situations are validated and logged.
4. Use guard clauses instead of nesting.
5. Protect Excel application state.
6. Always check objects before use.
7. Always validate external data.
8. Never use silent failures.
9. Log handled failures once, with context.
10. Design for failure, not success.
