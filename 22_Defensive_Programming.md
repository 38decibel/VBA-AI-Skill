# 22 - Defensive Programming

## Objective

Defensive programming is about building VBA code that:

- fails safely instead of crashing
- validates inputs early
- assumes external data is unreliable
- prevents silent corruption
- improves traceability and robustness

In VBA, where type safety and compiler guarantees are limited, defensive programming is **mandatory, not optional**.

---

# Golden Rule

> Never trust inputs: Excel data, user input, files, or external systems are all potentially invalid.

---

# 1. Validate Early, Fail Fast

## Principle

All validation must happen at the **entry point of a procedure**.

---

## Example

```vb
If ws Is Nothing Then Exit Sub
If Len(orderId) = 0 Then Exit Sub
If qty <= 0 Then Exit Sub
```

---

# 2. Null / Nothing Safety

## Always check objects

```vb
If lo Is Nothing Then Exit Sub
If rng Is Nothing Then Exit Sub
If wb Is Nothing Then Exit Sub
```

---

## ListObject special case

```vb
If lo.DataBodyRange Is Nothing Then Exit Sub
```

---

# 3. Range Safety Rules

Excel ranges can be empty, invalid, or unexpected.

## Safe pattern

```vb
If Target Is Nothing Then Exit Sub
If Target.CountLarge > 1 Then Exit Sub
```

---

# 4. Type Safety (Variant danger control)

VBA is loosely typed — defensive checks are required.

## Example

```vb
If Not IsNumeric(value) Then Exit Sub
If Not IsDate(value) Then Exit Sub
```

---

# 5. Error Handling is Mandatory

Every procedure must include structured error handling.

## Standard pattern

```vb
On Error GoTo ErrorHandler

' logic here

Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "ModuleName.ProcedureName"
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

Only in controlled blocks:

```vb
On Error Resume Next
Set wb = Workbooks.Open(path)
On Error GoTo 0
```

---

# 7. Defensive COM Handling

COM objects must always be protected:

```vb
If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
End If
```

---

# 8. File System Safety

Always validate:

```vb
If Not fso.FileExists(path) Then Exit Sub
If Not fso.FolderExists(folder) Then Exit Sub
```

---

# 9. Array Safety Rules

Excel arrays must be validated before use:

```vb
If IsEmpty(data) Then Exit Sub
If Not IsArray(data) Then Exit Sub
```

---

## Bounds safety

```vb
If UBound(data, 1) < 1 Then Exit Sub
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

without validation.

---

# 11. User Input Safety

All UserForm inputs must be validated:

## Required fields

```vb
If txtOrderId.Value = "" Then Exit Sub
```

## Numeric fields

```vb
If Not IsNumeric(txtQty.Value) Then Exit Sub
```

---

# 12. Excel State Protection

Always protect Excel environment state:

```vb
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
```

And restore:

```vb
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
```

---

# 13. Safe Exit Pattern

Use centralized cleanup:

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

```vb
If Not condition1 Then Exit Sub
If Not condition2 Then Exit Sub
If Not condition3 Then Exit Sub
```

---

# 15. Defensive Logging

Every failure must be traceable:

```vb
Utils_Log.Error Err, "Order.Process"
```

Include context:

```vb
Utils_Log.Error Err, "Order.Process - OrderId=" & orderId
```

---

# 16. Silent Failure is Forbidden

Never swallow errors:

Bad:

```vb
On Error Resume Next
```

without handling.

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

Each procedure must follow:

```
Validate → Process → Validate → Output → Log
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

Everything from Excel is Variant → validate.

---

## 21.4 External systems

COM, files, APIs → always fail-prone.

---

# 22. Defensive vs Overengineering

## Good defensive programming

- validation
- error handling
- safe defaults

## Bad overengineering

- excessive checks everywhere
- redundant validation
- performance degradation

---

# 23. AI Rules

When generating VBA code, AI must:

- always validate inputs
- always check objects for Nothing
- always implement error handling
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
    Utils_Log.Debug "Missing key: " & key
End If
```

---

# Golden Rules

1. Never trust inputs.
2. Validate early and often.
3. Always handle errors explicitly.
4. Use guard clauses instead of nesting.
5. Protect Excel application state.
6. Always check objects before use.
7. Always validate external data.
8. Never use silent failures.
9. Log all failures with context.
10. Design for failure, not success.