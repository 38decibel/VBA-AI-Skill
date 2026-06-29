# 14 - Arrays, Collections & Dictionaries

## Objective

This chapter defines how to handle in-memory data structures in VBA.

Efficient use of arrays, collections, and dictionaries is critical for:

- performance
- scalability
- clean architecture
- reducing Excel object interactions

VBA code must prioritize **in-memory processing over worksheet operations**.

---

# Golden Rule

> Always process data in memory first. Excel is only a source or destination, not a processing engine.

---

# 1. Arrays (Primary Structure for Performance)

## When to use arrays

Use arrays when:

- handling large datasets
- reading from ListObjects or Ranges
- transforming data
- avoiding repeated Excel calls

---

## 1.1 Preferred pattern

```vb
Dim data As Variant
data = lo.DataBodyRange.Value2
```

This creates a **2D array (1-based)**.

---

## 1.2 Array structure from Excel

Excel ranges produce:

```
(data(row, column))
```

Example:

```vb
data(1, 1) ' first row, first column
```

---

## 1.3 Writing back to Excel

```vb
lo.DataBodyRange.Value2 = data
```

---

## 1.4 Array rules

- always assume 2D arrays for Excel data
- never assume 1D arrays from Range
- always use `Value2`
- avoid ReDim inside loops

---

## 1.5 Looping arrays (correct)

```vb
Dim i As Long, j As Long

For i = LBound(data, 1) To UBound(data, 1)
    For j = LBound(data, 2) To UBound(data, 2)
        Debug.Print data(i, j)
    Next j
Next i
```

---

## 1.6 Avoid cell-by-cell logic

Bad:

```vb
ws.Cells(i, 1).Value = data(i)
```

Good:

```vb
ws.Range("A1:A" & lastRow).Value2 = data
```

---

# 2. Dynamic Arrays (ReDim / Preserve)

## Use sparingly

Dynamic arrays are allowed but should be minimized.

---

## 2.1 Basic usage

```vb
Dim arr() As String
ReDim arr(1 To 10)
```

---

## 2.2 Expanding arrays

```vb
ReDim Preserve arr(1 To newSize)
```

---

## ⚠️ Performance warning

`ReDim Preserve` is expensive.

Avoid in loops.

Bad:

```vb
For i = 1 To 10000
    ReDim Preserve arr(i)
Next i
```

---

## 2.3 Preferred alternative

Use:

- Dictionary
- pre-sized arrays
- Collection (limited cases)

---

# 3. Collections

## 3.1 When to use Collection

Use `Collection` when:

- order matters
- simple storage is needed
- no key lookup required

---

## 3.2 Example

```vb
Dim col As Collection
Set col = New Collection

col.Add "Item1"
col.Add "Item2"
```

---

## 3.3 Access

```vb
Debug.Print col(1)
```

---

## 3.4 Limitations

- no key overwrite control
- no existence check
- slower than Dictionary for lookups

---

## 3.5 Rule

> Collections are for simple lists only. Not for structured data.

---

# 4. Dictionary (Preferred Key-Value Structure)

## 4.1 Why Dictionary is preferred

- fast lookup
- key-based access
- prevents duplicates
- ideal for caching and mapping

---

## 4.2 Initialization

```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

---

## 4.3 Add values

```vb
dict.Add "A123", "ArticleName"
```

---

## 4.4 Access values

```vb
Debug.Print dict("A123")
```

---

## 4.5 Check existence

```vb
If dict.Exists("A123") Then
```

---

## 4.6 Update values

```vb
dict("A123") = "NewValue"
```

---

## 4.7 Dictionary rules

- keys must be unique
- always check `.Exists` before access
- prefer String keys
- avoid Variant keys

---

# 5. Arrays vs Collections vs Dictionaries

| Structure | Use case | Performance |
|----------|----------|-------------|
| Array | bulk data processing | ⭐⭐⭐⭐⭐ |
| Dictionary | key-value mapping | ⭐⭐⭐⭐⭐ |
| Collection | simple ordered list | ⭐⭐⭐ |

---

# 6. Recommended Architecture Pattern

## Step 1: Load from Excel

```vb
data = lo.DataBodyRange.Value2
```

## Step 2: Process in memory

```vb
dict("A123") = "ProcessedValue"
```

## Step 3: Write back if needed

```vb
lo.DataBodyRange.Value2 = data
```

---

# 7. Hybrid Pattern (Advanced)

Combine structures:

```vb
Dim data As Variant
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

data = lo.DataBodyRange.Value2
```

Use dictionary for:
- lookup
- enrichment
- deduplication

---

# 8. Common Pitfalls

## 8.1 Using arrays as dictionaries

Bad:

```vb
arr(key) = value
```

Unless indexed correctly → risky.

---

## 8.2 Repeated worksheet access

Bad:

```vb
ws.Cells(i, 1).Value
```

Always prefer arrays.

---

## 8.3 Missing bounds check

Always validate:

```vb
If Not dict.Exists(key) Then
```

---

## 8.4 Variant misuse

Avoid uncontrolled Variant propagation.

Always define structure early.

---

# 9. Memory Rules

- arrays are memory-efficient
- dictionaries consume more memory but are fast
- collections are lightweight but limited
- Excel object calls are the most expensive

---

# 10. Performance Rules

- minimize Excel interaction
- prefer bulk array processing
- avoid ReDim in loops
- avoid cell-by-cell operations
- prefer Dictionary for lookups

---

# 11. Debugging Structures

```vb
Utils_Log.Debug "Array", "Rows=" & UBound(data, 1)
Utils_Log.Debug "Dict", "Keys=" & dict.Count
```

---

# AI Rules

When generating VBA code, AI must:

- always prefer arrays for Excel data
- use `.Value2` when loading data
- avoid cell-by-cell processing
- use Dictionary for lookup logic
- avoid ReDim in loops
- always check `.Exists` before Dictionary access
- never use Collection for lookup logic
- ensure 2D array handling for Excel data
- optimize for in-memory processing
- minimize Excel object interactions

---

# Example

Bad:

```vb
For i = 1 To lastRow
    ws.Cells(i, 1).Value = ws.Cells(i, 2).Value
Next i
```

Good:

```vb
Dim data As Variant
data = ws.Range("A1:B1000").Value2
```

Better:

```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

---

# Golden Rules

1. Always process data in memory first.
2. Arrays are the default Excel data structure.
3. Dictionary is the default lookup structure.
4. Collections are for simple lists only.
5. Avoid cell-by-cell operations.
6. Avoid ReDim in loops.
7. Always use Value2 for Excel input.
8. Minimize Excel object calls.
9. Validate Dictionary access with Exists.
10. Optimize for performance and clarity.