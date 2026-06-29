# 05 - Error Handling

## Objective

Error handling is a fundamental part of application reliability.

Every VBA procedure should either:

- handle its own errors,
- propagate them to the caller,
- or explicitly document why no error handling is required.

Unexpected errors must **never** be ignored.

The project uses **`Utils_Log`** as the central logging service for diagnostics and troubleshooting.

---

# General Principles

Error handling should:

- prevent unexpected crashes
- keep Excel in a consistent state
- clean up resources
- provide enough information for troubleshooting
- log unexpected failures
- only notify the user when necessary

Error handling should **never** hide errors.

---

# Logging Policy

`Utils_Log` is the project's single source of truth for diagnostics.

Every unexpected error should be logged before exiting or propagating the error.

A log entry should contain, whenever possible:

- procedure name
- error number
- error description
- source
- workbook name
- worksheet name
- current row or record identifier
- additional business context

Logs should provide enough information to reproduce an issue without a debugger.

---

# Logging Levels

Use the appropriate logging level.

## Info

Normal execution.

Examples:

```vb
Utils_Log.Info "Import started"

Utils_Log.Info "153 records imported"
```

---

## Warning

Recoverable or unexpected business situations.

Examples:

```vb
Utils_Log.Warning "Unknown SAP material code"

Utils_Log.Warning "Label already exists"
```

---

## Error

Unexpected runtime failures.

Examples:

```vb
Utils_Log.Error _
    ProcedureName:="ExportLabels", _
    ErrObject:=Err
```

Every unexpected runtime error should generate an Error log.

---

# MsgBox Policy

`MsgBox` should **not** be the primary error reporting mechanism.

Use a `MsgBox` only when:

- the procedure is directly initiated by the user,
- the user must immediately react,
- continuing execution would be meaningless.

Never display technical information to the user if it has already been logged.

Preferred pattern:

```vb
Utils_Log.Error _
    ProcedureName:="ExportLabels", _
    ErrObject:=Err

MsgBox _
    "The operation could not be completed." & vbCrLf & _
    "See the application log for technical details.", _
    vbExclamation
```

Most helper modules should never display a `MsgBox`.

---

# Golden Rule

Never write:

```vb
On Error Resume Next
```

without:

- checking `Err.Number`
- validating the expected result
- restoring normal error handling immediately.

Silent failures create difficult bugs.

---

# Preferred Structure

Every non-trivial procedure should follow this structure.

```vb
Public Sub ExportLabels()

    On Error GoTo ErrorHandler

    ' Main code

CleanExit:

    Exit Sub

ErrorHandler:

    Utils_Log.Error _
        ProcedureName:="ExportLabels", _
        ErrObject:=Err

    Resume CleanExit

End Sub
```

---

# Functions

Functions follow the same philosophy.

```vb
Public Function FindLastRow(ws As Worksheet) As Long

    On Error GoTo ErrorHandler

    FindLastRow = ...

CleanExit:

    Exit Function

ErrorHandler:

    Utils_Log.Error _
        ProcedureName:="FindLastRow", _
        ErrObject:=Err

    FindLastRow = 0

    Resume CleanExit

End Function
```

---

# Clean Exit Section

Always centralize cleanup.

Typical cleanup includes:

```vb
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

Set rs = Nothing
Set cn = Nothing
```

Never duplicate cleanup code.

---

# Validation Before Execution

Prefer validating inputs instead of relying on runtime errors.

Examples:

```vb
If ws Is Nothing Then Exit Sub

If Len(filePath) = 0 Then Exit Sub

If Not FileExists(filePath) Then Exit Sub
```

Validation produces more predictable code.

---

# Expected Errors

Expected situations should be handled with validation.

Example:

```vb
If Not FileExists(filePath) Then

    Utils_Log.Warning _
        "Configuration file not found."

    Exit Sub

End If
```

Avoid using exceptions for normal business situations.

---

# Unexpected Errors

Unexpected runtime failures include:

- invalid object references
- COM failures
- printer unavailable
- worksheet deleted
- overflow
- external application failures

These errors should always be logged.

---

# On Error Resume Next

Allowed only for very small protected blocks.

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

Always restore standard error handling.

```vb
On Error GoTo 0
```

Never leave `Resume Next` active.

---

# Resource Cleanup

Always release external resources.

```vb
Set rs = Nothing
Set cn = Nothing
Set xlApp = Nothing
```

---

# Excel State

Whenever code modifies the Excel environment:

```vb
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
```

those settings must always be restored.

Even if an error occurs.

---

# Error Propagation

Some procedures should only log and rethrow.

Example:

```vb
Utils_Log.Error _
    ProcedureName:="LoadConfiguration", _
    ErrObject:=Err

Err.Raise _
    Err.Number, _
    Err.Source, _
    Err.Description
```

The caller decides how to react.

---

# Assertions

Validate assumptions early.

```vb
If ws Is Nothing Then

    Err.Raise _
        vbObjectError + 1000, _
        "LoadConfiguration", _
        "Worksheet cannot be Nothing."

End If
```

---

# Public Procedures

Public APIs should:

- validate inputs
- log unexpected errors
- leave Excel in a consistent state
- notify the user only when appropriate

---

# Private Helpers

Private helper procedures should:

- never display a `MsgBox`
- log errors
- propagate errors when appropriate

Helpers are infrastructure, not user interface.

---

# Debugging

During development:

```vb
Debug.Print Err.Number
Debug.Print Err.Description
```

For production code, prefer structured logging through `Utils_Log`.

---

# What NOT to Do

Never write:

```vb
On Error Resume Next

...

Exit Sub
```

Never ignore `Err.Number`.

Never swallow exceptions.

Never use empty error handlers.

Avoid:

```vb
ErrorHandler:

Resume Next
```

Never display a `MsgBox` from a utility module.

Never duplicate logging code throughout the application.

---

# AI Rules

When generating VBA code, AI should:

- always create structured error handlers for non-trivial procedures
- use `Utils_Log` for all unexpected errors
- avoid `MsgBox` except in UI workflows
- always restore Excel settings
- validate inputs before execution
- never silently ignore errors
- keep `On Error Resume Next` blocks as short as possible
- centralize cleanup in a `CleanExit` section
- propagate errors when appropriate
- produce logs containing meaningful diagnostic information

---

# Recommended Template

```vb
Public Sub Example()

    On Error GoTo ErrorHandler

    ' Main code

CleanExit:

    Exit Sub

ErrorHandler:

    Utils_Log.Error _
        ProcedureName:="Example", _
        ErrObject:=Err

    Resume CleanExit

End Sub
```

For UI procedures only:

```vb
ErrorHandler:

    Utils_Log.Error _
        ProcedureName:="ExportLabels", _
        ErrObject:=Err

    MsgBox _
        "The operation could not be completed." & vbCrLf & _
        "See the application log for details.", _
        vbExclamation

    Resume CleanExit
```

---

# Golden Rules

1. Never hide errors.
2. Always log unexpected failures following logging rules.
3. `MsgBox` is for users, not for developers.
4. Validate before executing.
5. Restore Excel state before exiting.
6. Clean up resources in a single location.
7. Keep `On Error Resume Next` blocks minimal.
8. Propagate errors when appropriate.
9. Make logs useful enough to diagnose issues without a debugger.
10. Robust applications fail gracefully and leave a clear diagnostic trail.