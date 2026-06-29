# 11 - Excel Object Model

## Objective

The Excel Object Model is the foundation of all VBA automation in this framework.

Incorrect handling of Excel objects leads to:

- performance issues
- unstable code
- memory leaks
- unpredictable results
- dependency on user context (ActiveSheet, Selection)

This chapter defines strict rules to ensure **explicit, safe, and performant Excel interactions**.

---

# Golden Rule

> Never rely on Excel's implicit context (ActiveSheet, ActiveWorkbook, Selection).

Always use explicit references.

---

# Excel Object Hierarchy

Understand and respect the hierarchy:

```
Application
  └── Workbooks
        └── Workbook
              └── Worksheets
                    └── Worksheet
                          └── Range
```

Every object must be explicitly referenced when possible.

---

# Worksheet Reference Strategy (IMPORTANT)

This framework distinguishes two types of worksheets:

## 1. UI / System Worksheets (CodeName mandatory)

These sheets are part of the application design and are owned by the VBA system.

Examples:
- configuration sheets
- export sheets
- log sheets
- dashboard / UI sheets
- control panels

### Rule

> UI/System worksheets must ALWAYS be referenced via **CodeName (VBName)**.

Example:

```vb
Dim ws As Worksheet
Set ws = wsConfig
```

```vb
Set ws = wsExport
Set ws = wsLog
```

### Forbidden for system sheets

```vb
Worksheets("Config")
Sheets("Export")
```

Reason:
- tab names can change
- CodeName is stable and design-time controlled
- prevents fragile dependencies

---

## 2. Data / External Worksheets (Name allowed)

These sheets are dynamic and not owned by the application.

Examples:
- imported SAP extracts
- user-generated sheets
- external files
- temporary datasets

### Rule

> Data worksheets must be referenced using `.Name` (string-based lookup) or dynamic references.

Example:

```vb
Set ws = wb.Worksheets(sheetNameFromUser)
Set ws = wb.Worksheets("SAP_EXPORT_2026")
```

---

## Summary Rule

| Type | Reference method |
|------|------------------|
| UI / System sheets | CodeName (VBName) |
| Data / External sheets | Worksheet.Name |

---

# Workbook Handling

## Always declare Workbooks explicitly

```vb
Dim wb As Workbook
Set wb = Workbooks("Orders.xlsx")
```

## Avoid ActiveWorkbook

Bad:

```vb
Set wb = ActiveWorkbook
```

---

# Worksheet Handling

## System sheets (preferred)

```vb
Set ws = wsConfig
Set ws = wsExport
```

## Data sheets (dynamic)

```vb
Set ws = wb.Worksheets("SAP_Export")
```

---

# Range Handling

## Always qualify Range with Worksheet

Good:

```vb
ws.Range("A1")
ws.Cells(1, 1)
```

Bad:

```vb
Range("A1")
Cells(1, 1)
```

---

# Cells vs Range

## Use Cells for dynamic access

```vb
ws.Cells(rowIndex, colIndex)
```

## Use Range for fixed references

```vb
ws.Range("A1:D10")
```

---

# Never Use Selection

Bad:

```vb
Selection.Copy
Selection.Value = "X"
```

Good:

```vb
ws.Range("A1").Value = "X"
```

---

# Avoid Activate / Select

Bad:

```vb
ws.Activate
ws.Range("A1").Select
```

Good:

```vb
ws.Range("A1").Value = "X"
```

---

# With Blocks

Use `With` only when it improves readability.

```vb
With ws
    .Cells(1, 1).Value = "Header"
End With
```

Avoid deep nesting.

---

# UsedRange is Dangerous

Avoid:

```vb
ws.UsedRange
```

Prefer explicit last row detection.

---

# Last Row Pattern

```vb
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
```

---

# Last Column Pattern

```vb
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
```

---

# Reading Data Efficiently

Bad (slow):

```vb
For i = 1 To 10000
    value = ws.Cells(i, 1).Value
Next i
```

Good (fast):

```vb
data = ws.Range("A1:A10000").Value
```

---

# Writing Data Efficiently

Bad:

```vb
ws.Cells(i, 1).Value = value
```

Good:

```vb
ws.Range("A1:A10000").Value = data
```

---

# Workbook Safety

Always validate objects:

```vb
If wb Is Nothing Then Exit Sub
If ws Is Nothing Then Exit Sub
```

---

# Object Lifecycle

Release objects when needed:

```vb
Set ws = Nothing
Set wb = Nothing
```

Especially for external workbooks.

---

# ScreenUpdating and Performance

```vb
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
```

Always restore:

```vb
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
```

---

# Hidden Dependencies (FORBIDDEN)

Never use:

- ActiveSheet
- ActiveWorkbook
- Selection
- ActiveCell

---

# Naming Rules

Use meaningful names:

Good:

```vb
wsOrders
wsConfig
wsExport
```

Bad:

```vb
ws1
ws2
ws3
```

---

# Debugging Tip

```vb
Utils_Log.Debug "Excel", "ActiveSheet=" & ActiveSheet.Name
```

Only in debug mode.

---

# AI Rules

When generating VBA code, AI must:

- always fully qualify Excel objects
- never use Selection / Activate / ActiveSheet
- use CodeName for UI/system sheets
- use string references for dynamic sheets
- prefer bulk operations
- validate objects before use
- minimize COM calls
- avoid UsedRange unless justified

---

# Example

Bad:

```vb
Range("A1").Select
Selection.Value = "Test"
```

Good:

```vb
ws.Cells(1, 1).Value = "Test"
```

---

# Golden Rules

1. Never rely on ActiveSheet or ActiveWorkbook.
2. Use CodeName for system/UI sheets.
3. Use Name for dynamic/external sheets.
4. Always qualify Range and Cells.
5. Avoid Selection and Activate.
6. Prefer bulk data operations.
7. Validate objects before use.
8. Minimize Excel COM calls.
9. Treat Excel as a strict object hierarchy.
10. Separate application sheets from data sheets.