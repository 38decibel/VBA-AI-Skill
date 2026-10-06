# 05 - Error Handling

## Objective

Error handling is a fundamental part of application reliability.

Every VBA procedure falls into exactly one of these categories:

- it **handles** errors (entry points, recoverable situations),
- it **cleans up and propagates** errors (procedures that own a resource or an application state),
- or it **lets errors propagate** untouched (everything else).

Unexpected errors must **never** be ignored, and must be logged **once**, at the boundary where they are finally handled.

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

A handler that does nothing except log and re-raise at every level of the call stack is noise:
it duplicates log entries and adds no value. Place handlers where they do real work.

---

# Procedure Roles (Where Handlers Belong)

| Role | Handler | What the handler does |
|------|---------|-----------------------|
| **Entry point** (macro, button, `Workbook_Open`, `Worksheet_Change`, form event) | Mandatory | Log once with `Utils_Log.Error`, tell the user when appropriate, `Resume CleanExit` |
| **Owns a resource or application state** (Excel settings, COM object, SAP session, file handle) | Mandatory | Clean up in `CleanExit`, then **re-raise** the original error to the caller (no logging unless adding context) |
| **Can recover** from a specific error | Local | Handle that error, re-raise everything else |
| **Plain helper** (pure logic, no resource, no recovery) | None | Let errors propagate; validate inputs with guard clauses |

Rules:

- Never write a handler whose only content is "log and re-raise" in a procedure that owns nothing to clean up.
- Never log the same failure at several levels of the call stack.
- When a procedure documents "no error handling needed", the reason is the role above (plain helper).

---

# Logging Policy

`Utils_Log` is the project's single source of truth for diagnostics.

Every unexpected error is logged **once**, by the procedure that finally handles it (normally the entry point),
before notifying the user.

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

Use the appropriate logging level. The canonical signatures are defined in chapter 10.

## Info

Normal execution.

```vb
Utils_Log.Info "ImportOrders", "Import started"
Utils_Log.Info "ImportOrders", "153 records imported"
```

---

## Warning

Recoverable or unexpected business situations.

```vb
Utils_Log.Warning "ImportOrders", "Unknown SAP material code", "Article=ABC123"
Utils_Log.Warning "PrintLabels", "Label already exists"
```

---

## Error

Unexpected runtime failures.

```vb
Utils_Log.Error Err, "ExportLabels"
```

Every unexpected runtime error that is **handled** generates exactly one Error log.

---

# Programming Errors vs User/Data Errors

Two kinds of problems exist and they are treated differently.

| Kind | Examples | Treatment |
|------|----------|-----------|
| **Programming error** (the caller broke the contract) | `Nothing` argument, empty required string, negative quantity where impossible, undefined `Enum` value | **Raise a custom error** (guard clause). Fail early and loudly |
| **User / data error** (expected in real life) | File not found, SAP material without price, empty cell, cancelled dialog, bad user input | **Validate**, log a `Warning`, give the user a clear and specific message, exit cleanly |

User-facing messages must be specific and actionable ("SKU 1234 has no sales price"),
never a raw technical message such as "Division by zero".

---

# Guard Clauses That Raise

A guard clause validates a parameter as one of the first executable statements and **raises** an error
when the contract is violated. It must never silently `Exit Sub` for a programming error:
a silent exit hides the bug and moves the failure far away from its cause.

Use the `Utils_Guard` module and the central `AppError` enumeration.

```vb
Option Explicit
Option Private Module

' Custom error numbers: keep ALL of them in this single enumeration to avoid overlaps.
Public Enum AppError
    ErrInvalidArgument = vbObjectError + 1000
    ErrNothingArgument
    ErrEmptyArgument
    ErrNegativeArgument
    ErrSapSessionUnavailable
End Enum

Public Sub NotNothing(ByVal value As Object, ByVal argumentName As String)

    If value Is Nothing Then
        Err.Raise AppError.ErrNothingArgument, "Utils_Guard.NotNothing", argumentName & " cannot be Nothing."
    End If

End Sub

Public Sub NotEmpty(ByVal value As String, ByVal argumentName As String)

    If Len(value) = 0 Then
        Err.Raise AppError.ErrEmptyArgument, "Utils_Guard.NotEmpty", argumentName & " cannot be empty."
    End If

End Sub

Public Sub NotNegative(ByVal value As Double, ByVal argumentName As String)

    If value < 0 Then
        Err.Raise AppError.ErrNegativeArgument, "Utils_Guard.NotNegative", argumentName & " cannot be negative."
    End If

End Sub
```

Usage:

```vb
Public Sub ExportWorksheet(ByVal ws As Worksheet, ByVal outputFolder As String)

    Utils_Guard.NotNothing ws, "ws"
    Utils_Guard.NotEmpty outputFolder, "outputFolder"

    ' ...

End Sub
```

Never compare `Err.Number` with a hard-coded integer: compare it with an `AppError` member.

When a value is **expected** to be missing (user data, configuration, optional file), do not raise:
validate, log a `Warning` and exit as described in "Expected Errors".

---

# MsgBox Policy

`MsgBox` should **not** be the primary error reporting mechanism.

Use a `MsgBox` only when:

- the procedure is directly initiated by the user,
- the user must immediately react,
- continuing execution would be meaningless.

Never display technical information to the user if it has already been logged.

Preferred pattern for an entry point:

```vb
Utils_Log.Error Err, "ExportLabels"

MsgBox _
    "The operation could not be completed." & vbCrLf & _
    "See the application log for technical details.", _
    vbExclamation
```

Most helper modules should never display a `MsgBox`.

---

# Standard Structure

Labels are always named `CleanExit` and `CleanFail`. Do not invent other names.

## Entry point (log and resume)

```vb
Public Sub ExportLabels()

    On Error GoTo CleanFail

    ' Main code

CleanExit:
    ' Cleanup code (runs with or without error)
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "ExportLabels"
    Resume CleanExit

End Sub
```

`Resume CleanExit` resets the error state and sends execution to the single exit point.
The nominal path must never run while an error is being handled, so every error path ends with a `Resume`.

## Procedure that owns state (clean up, then re-raise)

Raising an error from inside `CleanFail` would skip the cleanup. Capture the error, resume to `CleanExit`,
clean up, and only then re-raise the original error.

```vb
Public Sub LoadOrders(ByVal ws As Worksheet)

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    ' Main code

CleanExit:
    Application.ScreenUpdating = True

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If

    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit

End Sub
```

## Function

```vb
Public Function FindLastRow(ByVal ws As Worksheet) As Long

    Utils_Guard.NotNothing ws, "ws"

    FindLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

End Function
```

A plain helper such as this one has no handler.
When a function owns a resource, use the same structure as above and return a safe value or re-raise.

---

# Clean Exit Section

Always centralize cleanup in `CleanExit`.

Typical cleanup includes:

```vb
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

Set rs = Nothing
Set cn = Nothing
```

Never duplicate cleanup code, and never release a resource just before `Exit Sub` outside `CleanExit`:
an error would skip that release.

---

# Validation Before Execution

Prefer validating inputs instead of relying on runtime errors.

- Contract violations (programming errors): guard clause that **raises** (see above).
- Expected situations (user/data errors): validate and exit with a `Warning`.

---

# Expected Errors

Expected situations are handled with validation, not with exceptions.

```vb
If Not Utils_File.FileExists(filePath) Then

    Utils_Log.Warning "LoadConfiguration", "Configuration file not found.", "Path=" & filePath
    Exit Sub

End If
```

Avoid using the error mechanism for normal business situations.

---

# Unexpected Errors

Unexpected runtime failures include:

- invalid object references
- COM failures
- printer unavailable
- worksheet deleted
- overflow
- external application failures

These errors must always be logged by the procedure that finally handles them.

---

# On Error Resume Next

Allowed only for very small protected blocks, never to "make an error go away".

The best form is a dedicated `Try` function (see next section) where the `Resume Next` lives in a tiny
procedure of its own.

Inline form inside a procedure that has a handler:

```vb
On Error Resume Next
Set ws = wb.Worksheets("Config")
On Error GoTo CleanFail

If ws Is Nothing Then
    Utils_Log.Warning "LoadConfiguration", "Config worksheet not found"
    Exit Sub
End If
```

The protected block must be as short as possible and must be followed by a validation of its result.

---

# Restore Error Handling

After an inline `On Error Resume Next`, restore the **procedure's own handler**:

```vb
On Error GoTo CleanFail
```

Never use `On Error GoTo 0` in a procedure that has a handler: it disables that handler, and every later
error in the procedure becomes unhandled.
`On Error GoTo 0` is only acceptable inside a tiny `Try` function that has no handler of its own.

Never leave `Resume Next` active.

---

# The Try Pattern

Wrap an *expected* failure in a function that returns `Boolean` and delivers its result through a
`ByRef` output parameter prefixed with `out`.

```vb
Public Function TryGetWorksheet( _
    ByVal wb As Workbook, _
    ByVal sheetName As String, _
    ByRef outWorksheet As Worksheet) As Boolean

    Set outWorksheet = Nothing

    On Error Resume Next
    Set outWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    TryGetWorksheet = Not outWorksheet Is Nothing

End Function
```

Call site, with a named argument to make the output obvious:

```vb
Dim wsOrders As Worksheet

If Not Utils_Excel.TryGetWorksheet(wb, "Orders", outWorksheet:=wsOrders) Then
    Utils_Log.Warning "ImportOrders", "Orders worksheet not found"
    Exit Sub
End If
```

The `Try` function swallows only the one expected error and keeps the calling procedure free of
`Resume Next` blocks.

---

# Resource Cleanup

Always release external resources in `CleanExit`.

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

those settings must always be restored in `CleanExit`, even if an error occurs.

---

# Error Propagation

When a procedure must let an error continue to its caller after cleanup, use the "clean up, then re-raise"
structure above. The original number, source and description are preserved.

Do not wrap a propagated error in a new generic error: the caller needs the original information.

---

# Public Procedures

Public APIs should:

- validate inputs with guard clauses
- leave Excel in a consistent state
- notify the user only when they are entry points and it is appropriate

---

# Private Helpers

Private helper procedures should:

- never display a `MsgBox`
- not log errors they do not handle
- propagate errors

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

Never write a handler that only does `Resume Next`.

Never display a `MsgBox` from a utility module.

Never log the same error at several levels.

Never use error handling as normal control flow: check preconditions instead
(for example test that a file exists before opening it, and still handle the unexpected I/O failure).

Never raise a new error from inside a handler unless the intent is to propagate the original error.

---

# AI Rules

When generating VBA code, AI should:

- put a handler in entry points and in procedures that own a resource or application state
- not put a log-and-re-raise handler in plain helpers
- use `CleanExit` / `CleanFail` labels exclusively
- end every error path with `Resume CleanExit`
- clean up before re-raising, never after
- raise a custom error (`AppError` + `Utils_Guard`) for programming errors instead of a silent `Exit Sub`
- validate and log a `Warning` for expected user/data situations
- use `Utils_Log` once, where the error is finally handled
- avoid `MsgBox` except in UI workflows
- always restore Excel settings
- keep `On Error Resume Next` blocks as short as possible, preferably inside a `Try` function
- restore the procedure's own handler after an inline `Resume Next`, never `On Error GoTo 0`
- produce logs containing meaningful diagnostic information

---

# Golden Rules

1. Never hide errors.
2. Log an unexpected failure once, where it is finally handled.
3. `MsgBox` is for users, not for developers.
4. Programming errors raise; user and data errors are validated.
5. Restore Excel state before exiting.
6. Clean up resources in a single location (`CleanExit`).
7. Keep `On Error Resume Next` blocks minimal, preferably in `Try` functions.
8. Put handlers where they do real work.
9. Make logs useful enough to diagnose issues without a debugger.
10. Robust applications fail gracefully and leave a clear diagnostic trail.
