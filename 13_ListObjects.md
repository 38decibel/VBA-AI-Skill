# 13 - ListObjects (Excel Tables)

## Objective

Excel Tables (`ListObjects`) are the **preferred structure for structured data** in this framework.

They provide:

- stable column references
- automatic range expansion
- structured data handling
- safer alternatives to raw Range logic

However, improper usage can still lead to fragile or slow code.

This chapter defines strict rules for safe ListObject usage.

---

# Golden Rule

> Always use structured references (ListObject + column names), never positional assumptions.

---

# 1. Prefer ListObjects over Ranges for Data Tables

## Why ListObjects?

Compared to raw ranges:

- automatic resizing
- named columns
- safer iteration
- better readability
- reduced dependency on cell addresses

---

# 2. Always Use Explicit Table References

## Correct

```vb
Dim lo As ListObject
Set lo = wsOrders.ListObjects("tblOrders")
```

## Incorrect

```vb
Set lo = wsOrders.ListObjects(1)
```

Reason:
- index-based access is fragile
- table order can change

---

# 3. Never Rely on ActiveSheet Tables

Bad:

```vb
Set lo = ActiveSheet.ListObjects("tblOrders")
```

Good:

```vb
Set lo = wsOrders.ListObjects("tblOrders")
```

---

# 4. Always Use Named Tables

Each table MUST have a meaningful name:

Examples:

- `tblOrders`
- `tblConfig`
- `tblExport`
- `tblCustomers`

Avoid:

- Table1
- Table2
- default Excel naming

---

# 5. Access Data Using ListColumns (NOT Index)

## Correct

```vb
lo.ListColumns("ArticleCode").DataBodyRange
```

## Incorrect

```vb
lo.ListColumns(1).DataBodyRange
```

Reason:
- column order can change
- fragile automation

---

# 6. Reading Table Data Efficiently (VALUE2 RULE)

## Preferred (bulk read)

```vb
Dim data As Variant
data = lo.DataBodyRange.Value2
```

## Preferred (bulk write)

```vb
lo.DataBodyRange.Value2 = data
```

---

## Why Value2 is mandatory for ListObjects

- faster than `.Value`
- no Excel formatting interpretation
- consistent raw data handling
- stable for ETL / transformation logic
- avoids implicit type conversions

---

## When NOT to use Value2

Only use `.Value` when Excel-specific semantics are required:

- formatted display values
- UI interaction logic
- legacy compatibility cases

Otherwise:

> `.Value2` is the default standard.

---

# 7. Writing to Tables

## Bulk write (preferred)

```vb
lo.DataBodyRange.Value2 = data
```

## Row-by-row (avoid when possible)

```vb
lo.ListRows(i).Range.Cells(1, lo.ListColumns("ArticleCode").Index).Value = articleCode
```

Use row-by-row only when necessary.

---

# 8. Adding Rows

## Correct

```vb
Dim newRow As ListRow
Set newRow = lo.ListRows.Add
```

Then populate:

```vb
newRow.Range.Cells(1, lo.ListColumns("ArticleCode").Index).Value = articleCode
```

---

# 9. Avoid Resize Hacks on Tables

Bad:

```vb
lo.Resize ws.Range("A1:D1000")
```

Reason:
- breaks table integrity
- can corrupt structure
- loses structured behavior

Prefer adding/removing rows instead.

---

# 10. Safe Column Access Pattern

Preferred method:

```vb
Dim colIndex As Long
colIndex = lo.ListColumns("ArticleCode").Index
```

Then use:

```vb
data = lo.DataBodyRange.Value2
```

Process in memory instead of repeated access.

---

# 11. Empty Table Handling

Always validate table state:

```vb
If lo.DataBodyRange Is Nothing Then Exit Sub
```

Empty tables return `Nothing`, not an empty array.

---

# 12. Looping Through Table Rows

## Correct

```vb
Dim lr As ListRow

For Each lr In lo.ListRows
    Debug.Print lr.Range.Cells(1, 1).Value
Next lr
```

## Better (performance)

```vb
Dim data As Variant
data = lo.DataBodyRange.Value2
```

Process in memory instead of row iteration.

---

# 13. Table Header Usage

Always prefer headers over positions:

```vb
lo.ListColumns("Status")
```

Never:

```vb
lo.ListColumns(3)
```

---

# 14. Avoid Mixed Range + Table Logic

Bad:

```vb
ws.Range("A2:A1000")
lo.ListColumns("A")
```

Good:

```vb
lo.DataBodyRange
```

Rule:

> Once a dataset is a ListObject, never revert to raw Range logic.

---

# 15. Table Creation (Controlled Use)

Tables should be created only in initialization or import layers.

```vb
Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
lo.Name = "tblOrders"
```

---

# 16. Naming Rules (Critical)

## Table names must:

- be stable
- describe business entity
- use prefix `tbl`

Examples:

- `tblOrders`
- `tblInvoices`
- `tblConfig`
- `tblExportData`

---

# 17. Performance Rules

- always use `.Value2` for bulk operations
- avoid cell-by-cell access
- minimize ListRow iteration
- avoid repeated ListColumns lookups inside loops
- process data in memory whenever possible

---

# 18. Common Pitfalls

## Empty table crash

```vb
lo.DataBodyRange.Value2
```

→ fails if table has no rows

## Column name mismatch

```vb
lo.ListColumns("Article")
```

→ breaks if renamed

## Table not found

Always validate existence.

---

# 19. Integration with CodeName Rule

System sheets:

```vb
wsOrders.ListObjects("tblOrders")
```

Never:

```vb
Worksheets("Orders").ListObjects("tblOrders")
```

---

# 20. Debugging Tables

Use logging instead of manual inspection:

```vb
Utils_Log.Debug "Table", "Rows=" & lo.ListRows.Count
Utils_Log.Debug "Table", "Cols=" & lo.ListColumns.Count
```

---

# AI Rules

When generating VBA code, AI must:

- always use ListObjects for structured data
- always use `.Value2` for data operations
- never rely on column indexes directly
- always use ListColumns("Name")
- prefer bulk DataBodyRange operations
- avoid row-by-row loops when possible
- validate empty tables safely
- ensure stable table naming
- avoid mixing Range and ListObject logic
- ensure performance-aware design
- assume tables may be empty

---

# Example

Bad:

```vb
ws.Cells(i, 1).Value = ws.Cells(i, 2).Value
```

Good:

```vb
Dim data As Variant
data = lo.DataBodyRange.Value2
```

---

# Golden Rules

1. Always use ListObjects for structured data.
2. Always use `.Value2` for data operations.
3. Never rely on column indexes directly.
4. Always use named tables.
5. Prefer bulk operations over loops.
6. Never mix Range and ListObject logic.
7. Validate empty tables safely.
8. Avoid row-by-row processing when possible.
9. Tables are business entities, not cell ranges.
10. Treat ListObjects as the default data model in Excel.