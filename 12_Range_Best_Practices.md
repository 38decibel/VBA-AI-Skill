# 12 - Range Best Practices

## Objective

The `Range` object is one of the most powerful and dangerous parts of the Excel Object Model.

Incorrect usage leads to:

- poor performance
- unstable macros
- hidden bugs
- excessive memory usage
- dependency on ActiveSheet (critical anti-pattern)

This chapter defines strict rules for safe and efficient Range usage.

---

# Golden Rule

> Always fully qualify ranges and avoid implicit worksheet context.

---

# 1. Always Qualify Ranges

## Correct

```vb
ws.Cells(1, 1).Value = "Header"
ws.Range("A1").Value = "Header"
```

## Incorrect

```vb
Range("A1").Value = "Header"
Cells(1, 1).Value = "Header"
```

These depend on `ActiveSheet`.

---

# 2. Never Use Selection

## Bad

```vb
Range("A1").Select
Selection.Value = "Test"
```

## Good

```vb
ws.Range("A1").Value = "Test"
```

Selection introduces:

- UI dependency
- performance overhead
- non-repeatable behavior

---

# 3. Avoid Entire Column References

## Bad

```vb
ws.Range("A:A")
ws.Columns("A")
```

Problems:
- extremely slow on large sheets
- includes unused cells
- unpredictable UsedRange expansion

## Good

```vb
ws.Range("A1:A" & lastRow)
```

---

# 4. Always Use Dynamic Last Row / Column

## Last Row

```vb
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
```

## Last Column

```vb
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
```

---

# 5. Prefer Bulk Read / Write

## Bad (slow)

```vb
For i = 1 To lastRow
    ws.Cells(i, 1).Value = data(i)
Next i
```

## Good (fast)

```vb
ws.Range("A1:A" & lastRow).Value = data
```

---

# 6. Use Variant Arrays for Performance

When reading large datasets:

```vb
Dim data As Variant
data = ws.Range("A1:D10000").Value
```

Then process in memory instead of repeated worksheet access.

---

# 7. Avoid Repeated Range Calls

## Bad

```vb
ws.Range("A1").Value = "X"
ws.Range("A1").Font.Bold = True
ws.Range("A1").Interior.Color = vbRed
```

## Good

```vb
With ws.Range("A1")
    .Value = "X"
    .Font.Bold = True
    .Interior.Color = vbRed
End With
```

---

# 8. Avoid Merged Cells

Merged cells cause:

- broken navigation
- incorrect row/column calculations
- unstable automation

Rule:

> Never rely on merged cells in automated sheets.

Prefer:

- Center Across Selection (formatting alternative)
- structured layout design

---

# 9. Use Named Ranges Only for Stable References

Named ranges are acceptable when:

- they are static
- defined in configuration sheets
- not dynamically modified

Example:

```vb
ws.Range("Config_Path").Value
```

Avoid dynamic named ranges for large datasets.

---

# 10. Use Resize and Offset Carefully

## Resize

```vb
ws.Range("A1").Resize(lastRow, 3)
```

## Offset

```vb
ws.Range("A1").Offset(1, 0)
```

Rules:

- always ensure bounds are controlled
- avoid chaining without clarity

---

# 11. Avoid Volatile Range Patterns

Bad patterns:

```vb
ws.UsedRange
ws.Cells.SpecialCells(xlCellTypeVisible)
```

Why:
- unpredictable
- slow
- depends on Excel internal state

Prefer explicit boundaries.

---

# 12. Range and CodeName Rule Integration

For system worksheets:

```vb
wsConfig.Range("A1").Value = "X"
```

Never:

```vb
Worksheets("Config").Range("A1")
```

For data sheets:

```vb
wb.Worksheets(sheetName).Range("A1")
```

---

# 13. Use With Blocks for Range Groups

Good:

```vb
With ws.Range("A1:D1")
    .Font.Bold = True
    .Interior.Color = vbYellow
End With
```

Avoid deep nesting of With blocks.

---

# 14. Avoid Hardcoded Addresses in Business Logic

Bad:

```vb
ws.Range("A1:D500")
```

Good:

```vb
ws.Range("A1:D" & lastRow)
```

Even better:

```vb
ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
```

---

# 15. Safe Range Construction Pattern

Preferred universal pattern:

```vb
Dim rng As Range

Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
```

This avoids string-based addressing entirely.

---

# 16. Range Validation

Always validate boundaries:

```vb
If lastRow < 1 Then Exit Sub
If lastCol < 1 Then Exit Sub
```

Avoid invalid range creation.

---

# 17. Intersection Usage (Advanced)

Use `Intersect` carefully:

```vb
If Not Intersect(Target, ws.Range("A:A")) Is Nothing Then
```

Rules:
- only in event-driven macros
- never in performance-critical loops

---

# 18. Performance Rules

- minimize Range calls inside loops
- prefer arrays over cell-by-cell access
- reduce COM interactions
- avoid volatile Range properties

---

# 19. Debugging Ranges

Use logging instead of inspecting manually:

```vb
Utils_Log.Debug "Range", "LastRow=" & lastRow
```

---

# AI Rules

When generating VBA code, AI must:

- always fully qualify Range and Cells
- avoid Selection and Activate
- prefer bulk operations over loops
- use arrays for large datasets
- avoid entire column references
- compute lastRow/lastCol explicitly
- avoid UsedRange unless justified
- prefer Range built via Cells objects
- ensure performance-aware design
- never assume ActiveSheet context

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

Best (structured range):

```vb
ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 3)).Value = data
```

---

# Golden Rules

1. Always fully qualify ranges.
2. Never use Selection.
3. Avoid entire column references.
4. Prefer bulk operations.
5. Use arrays for performance.
6. Always compute boundaries explicitly.
7. Avoid merged cells.
8. Prefer Cells-based range construction.
9. Minimize COM calls.
10. Treat Range as a high-cost object.