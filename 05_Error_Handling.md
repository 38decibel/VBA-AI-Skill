# 05 - Error Handling

## Objective

Error handling is not optional.

Every VBA procedure should either:

- handle its own errors,
- propagate them to the caller,
- or explicitly document why no error handling is required.

A program that crashes with a generic VBA message box is considered unfinished.

---

# General Principles

Error handling should:

- prevent unexpected crashes
- provide meaningful information
- leave Excel in a consistent state
- clean up resources
- help debugging

Error handling should **never** hide errors.

---

# Golden Rule

Never write:

```vb
On Error Resume Next
```

without immediately checking:

- `Err.Number`
- the expected result
- and restoring normal error handling.

Silent failures create difficult bugs.

---

# Preferred Structure

Every non-trivial procedure should follow this pattern:

```vb
Public Sub ExportData()

    On Error GoTo ErrorHandler

    ' Main code

CleanExit:

    Exit Sub

ErrorHandler:

    ' Cleanup

    MsgBox Err.Description

    Resume CleanExit

End Sub
```

---

# Functions

Functions follow the same structure.

```vb
Public Function FindLastRow(ws As Worksheet) As Long

    On Error GoTo ErrorHandler

    FindLastRow = ...

CleanExit:

    Exit Function

ErrorHandler:

    FindLastRow = 0

    Resume CleanExit

End Function
```

---

# Clean Exit Section

Always centralize cleanup.

Examples:

```vb
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

Set rs = Nothing
Set cn = Nothing
```

Never duplicate cleanup code.

---

# Error Handler Section

The error handler should:

- identify the procedure
- include the error number
- include the description
- clean resources
- optionally rethrow

Example:

```vb
MsgBox _
    "ExportData failed." & vbCrLf & _
    "Error " & Err.Number & vbCrLf & _
    Err.Description
```

---

# Expected Errors

Some errors are part of normal execution.

Example:

File not found.

Instead of crashing:

```vb
If Dir(filePath) = "" Then

    MsgBox "File not found."

    Exit Sub

End If
```

Prefer validation over exception handling.

---

# Unexpected Errors

Unexpected errors should be trapped.

Examples:

- invalid worksheet
- COM failures
- printer unavailable
- invalid object
- runtime overflow

---

# Avoid Nested Error Handling

Avoid:

```vb
On Error GoTo A

...

On Error GoTo B

...

On Error GoTo C
```

Keep one handler per procedure.

---

# On Error Resume Next

Allowed only for very specific operations.

Example:

```vb
On Error Resume Next

Set ws = Worksheets("Config")

If Err.Number <> 0 Then

    Err.Clear

End If

On Error GoTo 0
```

The protected block should be as short as possible.

---

# Restore Error Handling

Always restore immediately.

```vb
On Error GoTo 0
```

Never leave `Resume Next` active.

---

# Validation Before Action

Prefer:

```vb
If ws Is Nothing Then Exit Sub
```

instead of relying on runtime errors.

Likewise:

```vb
If Len(filePath) = 0 Then Exit Sub
```

---

# Resource Cleanup

Always release external resources.

Examples:

```vb
Set rs = Nothing
Set cn = Nothing
Set xlApp = Nothing
```

---

# Excel State

If code modifies Excel settings:

```vb
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
```

they **must** always be restored.

Even after an error.

---

# Logging

Complex applications should log errors.

Useful information:

- procedure name
- error number
- description
- date/time
- user
- workbook
- optional stack information

Example:

```
2026-06-15 14:23

ExportLabels

Error 91

Object variable not set
```

---

# Rethrowing Errors

Sometimes the caller should decide.

Example:

```vb
Err.Raise Err.Number, _
          "ExportData", _
          Err.Description
```

Do not swallow important errors.

---

# Assertions

Validate assumptions early.

Example:

```vb
If ws Is Nothing Then

    Err.Raise vbObjectError + 1, _
              "LoadConfiguration", _
              "Worksheet is required."

End If
```

---

# Avoid Generic Messages

Bad:

```vb
MsgBox "Error"
```

Good:

```vb
MsgBox _
"Unable to export labels." & vbCrLf & _
Err.Description
```

---

# Error Numbers

Always preserve:

```vb
Err.Number
Err.Description
Err.Source
```

Never replace useful information with generic text.

---

# Public APIs

Public procedures should fail gracefully.

They should:

- validate inputs
- report meaningful errors
- avoid leaving Excel unstable

---

# Private Helpers

Private helper functions may simply:

```vb
Err.Raise
```

and let the caller handle the error.

---

# Debugging

During development:

```vb
Debug.Print Err.Number
Debug.Print Err.Description
```

Avoid leaving debugging code in production unless intentionally logging.

---

# What NOT to Do

Never write:

```vb
On Error Resume Next

...

Exit Sub
```

Never ignore `Err.Number`.

Never suppress every error.

Never use empty error handlers.

Example to avoid:

```vb
ErrorHandler:

Resume Next
```

---

# AI Rules

When generating VBA code, AI should:

- always create a structured error handler for non-trivial procedures
- include a `CleanExit` section when cleanup is required
- restore Excel settings before exiting
- never leave `On Error Resume Next` active
- validate inputs before performing operations
- provide meaningful error messages
- preserve `Err.Number`, `Err.Description` and `Err.Source`
- never silently ignore errors

---

# Recommended Template

```vb
Public Sub Example()

    On Error GoTo ErrorHandler

    ' Main code

CleanExit:

    Exit Sub

ErrorHandler:

    MsgBox _
        "Example failed." & vbCrLf & _
        "Error " & Err.Number & vbCrLf & _
        Err.Description, vbCritical

    Resume CleanExit

End Sub
```

---

# Golden Rules

1. Never hide errors.
2. Validate before executing.
3. Use one error handler per procedure.
4. Always restore Excel state.
5. Always clean resources.
6. Keep `Resume Next` blocks as short as possible.
7. Report meaningful information.
8. Let callers handle errors when appropriate.
9. Never duplicate cleanup code.
10. Robust code fails gracefully.