# 10 - Logging

## Objective

Logging is a core infrastructure component of the VBA-AI-Skill framework.

It provides:

- execution traceability
- debugging capability (Immediate Window)
- production diagnostics (file output)
- audit trail for business processes

Without structured logging, complex VBA applications become difficult to debug and maintain.

---

# Central Logging System

The project uses a single logging module:

> **`Utils_Log`**

All logs must go through this module.

No direct `Debug.Print`, `MsgBox`, or ad-hoc logging is allowed in production code.

---

# Golden Rule

> Every important event must produce a structured log entry.

---

# Log Format (Core Principle)

All logs must follow a **single structured one-line format**:

```
Timestamp | Level | Caller | Message | Context
```

Example:

```
2026-06-29 14:22:11 | INFO  | ExportLabels | Start export
2026-06-29 14:22:12 | WARN  | ImportSAP    | Missing article code | Row=42
2026-06-29 14:22:13 | ERROR | ExportLabels | Object variable not set
```

This format must be identical regardless of output target.

---

# Output Channels

A log entry can be routed to multiple outputs.

## 1. Immediate Window (Development)

Used for:

- debugging
- execution tracing
- developer feedback

This is the primary developer view in VBA.

---

## 2. File Output (Optional / Production)

Used for:

- production debugging
- audit trail
- post-mortem analysis
- long-running processes

Example:

```vb
Print #fileNumber, formattedLogLine
```

File logging must reuse the same format as the Immediate Window.

---

# Logging Levels

## Info

Normal execution flow.

```vb
Utils_Log.Info "ExportLabels", "Start export"
Utils_Log.Info "ImportSAP", "152 rows processed"
```

---

## Warning

Non-critical issues or recoverable anomalies.

```vb
Utils_Log.Warning "ImportSAP", "Unknown article code", "Article=ABC123"
```

---

## Error

Unexpected runtime failures.

```vb
Utils_Log.Error Err, "ExportLabels"
```

Must include:
- error object (`Err`)
- caller procedure
- optional context

---

## Debug (Developer View)

### IMPORTANT CONCEPT

`Utils_Log.Debug` is the **primary developer view**, equivalent to `Debug.Print`, but controlled by a **global switch**.

It is used for development tracing only.

---

### Behavior

- If debug mode is **enabled** → output to Immediate Window
- If debug mode is **disabled** → no output (zero overhead)

---

### Usage

```vb
Utils_Log.Debug "ExportLabels", "Row=" & rowIndex
Utils_Log.Debug "ImportSAP", "Raw value=" & rawValue
```

---

### Global Switch

Debug output is controlled by a global flag:

```vb
Public gEnableDebug As Boolean
```

or equivalent configuration:

```vb
Public Const DEBUG_MODE As Boolean = True
```

---

### Rule

> Debug logs must never appear in production unless explicitly enabled.

---

### Design Principle

Debug logging is not a separate system.

It is:

> A filtered log level inside the same logging framework.

---

# Logging Rules

## 1. Single format everywhere

The same log format must be used for:

- Immediate Window
- File output
- future extensions (database, telemetry)

---

## 2. No silent failures

Never ignore errors without logging.

Bad:

```vb
On Error Resume Next
Set ws = Worksheets("Config")
```

Good:

```vb
On Error Resume Next
Set ws = Worksheets("Config")

If ws Is Nothing Then
    Utils_Log.Warning "Init", "Config worksheet not found"
End If

On Error GoTo 0
```

---

## 3. No duplicate logging

Each event should be logged once per layer.

---

## 4. Logging must never break execution

Logging is a non-critical infrastructure layer.

---

## 5. Avoid excessive logging in loops

Bad:

```vb
For i = 1 To 10000
    Utils_Log.Info "Loop", "Row " & i
Next i
```

Good:

```vb
If i Mod 500 = 0 Then
    Utils_Log.Info "Loop", "Processed " & i & " rows"
End If
```

---

# Error Logging Pattern

```vb
ErrorHandler:

    Utils_Log.Error Err, "ExportLabels"

    Resume CleanExit
```

---

# Context Logging

```vb
Utils_Log.Warning "ImportSAP", "Invalid article code", "Row=42"
```

or:

```vb
Utils_Log.Error Err, "ImportSAP", "Row=42, File=orders.xlsx"
```

---

# Business Event Logging

```vb
Utils_Log.Info "ImportSAP", "Start import"
Utils_Log.Info "BuildLabels", "Generating ZPL"
Utils_Log.Info "PrintService", "Sending to printer"
```

---

# Performance Considerations

- avoid logging in tight loops
- prefer aggregated logs
- disable Debug logs in production if needed
- keep string building lightweight

---

# Integration with Error Handling

Logging is integrated with error handling:

- error handling detects failures
- logging records them centrally via `Utils_Log`

---

# What NOT to Do

Never:

- use `MsgBox` for logging
- rely on `Debug.Print` directly
- mix log formats
- duplicate logs across layers
- log without context
- overload loops with logs
- break execution with logging

---

# AI Rules

When generating VBA code, AI must:

- always use `Utils_Log`
- enforce structured one-line logs
- include Caller + Message
- use correct log levels
- respect Debug switch behavior
- avoid logging in tight loops
- never use MsgBox for technical logging
- ensure logging does not affect execution
- ensure consistency across outputs

---

# Example

```vb
Public Sub ExportLabels()

    On Error GoTo ErrorHandler

    Utils_Log.Info "ExportLabels", "Start export"

    Utils_Log.Debug "ExportLabels", "Initializing variables"

    ' Main logic

CleanExit:

    Utils_Log.Info "ExportLabels", "Finished export"
    Exit Sub

ErrorHandler:

    Utils_Log.Error Err, "ExportLabels"

    Resume CleanExit

End Sub
```

---

# Golden Rules

1. One log format for all outputs.
2. Immediate Window is the primary developer view.
3. Debug logging is controlled by a global switch.
4. File logging is optional but consistent.
5. Every error must be logged.
6. No silent failures.
7. No logging inside hot loops.
8. Logs must be structured and readable.
9. Caller + Message are mandatory.
10. Logging is a diagnostic system, not noise.