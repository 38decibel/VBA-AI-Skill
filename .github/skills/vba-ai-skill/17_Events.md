# 17 - Events (Excel VBA Event Model)

## Objective

Excel Events are powerful but **dangerous if uncontrolled**.

They can easily lead to:

- infinite loops
- performance degradation
- unpredictable side effects
- hidden logic execution

This chapter defines strict rules for safe event handling.

---

# Golden Rule

> Events must only trigger orchestration logic â€” never business logic directly.

---

# 1. Types of Excel Events

## Worksheet Events
- `Worksheet_Change`
- `Worksheet_SelectionChange`
- `Worksheet_Activate`

## Workbook Events
- `Workbook_Open`
- `Workbook_BeforeClose`
- `Workbook_SheetChange`

## Application Events (advanced)
- `Application.OnTime`
- `Application events via class modules`

---

# 2. Event Responsibility Rule

Events must ONLY:

- detect a trigger
- validate minimal conditions
- call a controller / module
- log activity

---

## Forbidden in events

- business logic
- loops over large datasets
- Excel transformations
- direct ListObject processing
- heavy calculations

---

# 3. Standard Event Pattern

## Correct structure

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo CleanFail

    Utils_Log.Debug "Event", "Worksheet_Change triggered"

    EventController.HandleWorksheetChange Me, Target
CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Worksheet_Change"
    Resume CleanExit
End Sub
```

---

# 4. Prevent Infinite Loops

## Problem

Events triggered by code can re-trigger themselves.

Example:

```vb
Range("A1").Value = "X" ' triggers Worksheet_Change
```

---

## Solution: Disable Events

```vb
Application.EnableEvents = False

' safe modification
ws.Range("A1").Value = "X"

Application.EnableEvents = True
```

---

## Safe pattern

```vb
On Error GoTo CleanFail

Application.EnableEvents = False

' logic here

CleanExit:
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Event"
    Resume CleanExit
```

---

# 5. Event Filtering Rule

Never process entire sheet changes blindly.

## Bad

```vb
If Not Intersect(Target, ws.Range("A:Z")) Is Nothing Then
```

---

## Good

Limit scope:

```vb
If Not Intersect(Target, ws.Range("A1:A100")) Is Nothing Then
```

Even better:

```vb
If Target.CountLarge > 1 Then Exit Sub
```

---

# 6. Avoid Heavy Processing in Events

Bad:

```vb
For Each cell In Target
    ' complex logic
Next cell
```

Good:

```vb
EventController.HandleChange Target
```

---

# 7. Workbook_Open Pattern

```vb
Private Sub Workbook_Open()

    On Error GoTo CleanFail

    Utils_Log.Info "Workbook", "Application started"

    AppController.Initialize
CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Workbook_Open"
    Resume CleanExit
End Sub
```

---

# 8. Worksheet_Change Best Practice

## Minimal event handler

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo CleanFail

    If Target.CountLarge > 1 Then Exit Sub

    EventController.HandleCellChange Me, Target
CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Worksheet_Change"
    Resume CleanExit
End Sub
```

---

# 9. Event Controller Pattern

Events must delegate to controllers:

```
Event â†’ Controller Module â†’ Class Logic â†’ Excel
```

Example:

```vb
EventController.HandleWorksheetChange(ws, Target)
```

---

# 10. Event Reentrancy Protection

Use a global flag if needed:

```vb
Public IsEventRunning As Boolean
```

Usage:

```vb
If IsEventRunning Then Exit Sub
IsEventRunning = True

' logic

IsEventRunning = False
```

---

# 11. Application Events (Advanced)

Use class modules for Application-level events:

```vb
Private WithEvents App As Application
```

---

# 12. Event Logging Rule

Every event must be logged:

```vb
Utils_Log.Debug "Event", "Change detected in " & Target.Address
```

---

# 13. Error Handling in Events

Mandatory pattern:

```vb
On Error GoTo CleanFail

' logic

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "EventName"
    Resume CleanExit
```

An event handler is an entry point: it logs once and never lets the error escape to Excel.
If it changes application state (`EnableEvents`), the restore belongs in `CleanExit` (chapter 05).

Never allow silent failures.

---

# 14. No Business Logic in Events

Bad:

```vb
If Target.Value > 100 Then
    discount = 0.1
End If
```

Good:

```vb
PricingController.ApplyRules Target
```

---

# 15. Performance Rules

- avoid loops in events
- avoid Excel reads/writes inside events
- minimize triggered operations
- use arrays in controller layer only

---

# 16. Safe Target Handling

Always validate Target:

```vb
If Target Is Nothing Then Exit Sub
If Target.CountLarge > 1 Then Exit Sub
```

---

# 17. Events and ListObjects

Special care:

```vb
If Not Intersect(Target, lo.DataBodyRange) Is Nothing Then
```

Rules:

- always ensure DataBodyRange is not Nothing
- avoid full column monitoring
- prefer structured table events

---

# 18. Disable Events Rule

Only disable events when necessary:

- programmatic updates
- bulk writes
- controlled automation

Always restore state.

---

# 19. Debugging Events

Use logging instead of MsgBox:

```vb
Utils_Log.Debug "Event", "Triggered at " & Now
```

---

# 20. AI Rules

When generating event code, AI must:

- keep event procedures minimal
- always delegate to controllers
- never include business logic in events
- always implement error handling
- always use Utils_Log
- avoid loops in event handlers
- use EnableEvents carefully
- prevent reentrancy issues
- avoid full-sheet triggers
- ensure safe Target validation

---

# Example

## Bad

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    Target.Value = Target.Value * 2

End Sub
```

---

## Good

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo CleanFail

    If Target.CountLarge > 1 Then Exit Sub

    EventController.HandleWorksheetChange Me, Target
CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Worksheet_Change"
    Resume CleanExit
End Sub
```

---

# Golden Rules

1. Events must be thin.
2. No business logic in events.
3. Always delegate to controllers.
4. Always use error handling.
5. Always log event triggers.
6. Avoid loops in events.
7. Control EnableEvents carefully.
8. Prevent reentrancy issues.
9. Limit event scope strictly.
10. Treat events as triggers, not logic.
