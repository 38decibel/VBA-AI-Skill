# 07 - Data Types

## Objective

Choosing the correct data type is one of the most important aspects of writing robust and performant VBA code.

Bad type choices lead to:

- runtime errors
- performance issues
- memory overhead
- implicit conversions
- unpredictable behavior

Good type choices make code:

- safer
- faster
- easier to understand

---

# Golden Rule

Never use `Variant` by default.

Always choose the most specific type possible.

---

# Core VBA Data Types

## String

Use for text.

```vb
Dim customerName As String
Dim articleCode As String
```

Use `String` for all textual data, including numeric-like identifiers (SAP codes, barcodes, etc.).

---

## Long (preferred integer type)

Use for all whole numbers.

```vb
Dim rowIndex As Long
Dim quantity As Long
Dim lastRow As Long
```

### Why not Integer?

`Integer` is legacy in VBA and internally converted to `Long` anyway.

Always prefer:

```vb
Long
```

Never:

```vb
Integer
```

---

## Double

Use for decimals and calculations.

```vb
Dim price As Double
Dim weight As Double
Dim total As Double
```

Use when precision beyond whole numbers is required.

---

## Currency (financial precision)

Use for financial values when rounding errors must be avoided.

```vb
Dim unitPrice As Currency
Dim totalAmount As Currency
```

Preferred for:

- money
- accounting values
- invoices

Avoid `Double` for currency unless necessary.

---

## Boolean

Use for true/false logic.

```vb
Dim isValid As Boolean
Dim hasError As Boolean
Dim canExport As Boolean
```

Boolean variables must be named as questions (see naming rules).

---

## Date

Use for date/time values.

```vb
Dim creationDate As Date
Dim lastUpdated As Date
Dim exportTime As Date
```

VBA `Date` includes both date and time.

---

## Variant (restricted use)

Use only when absolutely necessary.

Allowed cases:

- unknown data source
- Excel Range `.Value` before casting
- API responses with dynamic structure
- legacy compatibility

Example:

```vb
Dim rawValue As Variant
rawValue = ws.Cells(1, 1).Value
```

### Mandatory rule

Every `Variant` must be:

- validated
- converted
- or explicitly justified

Never leave `Variant` flowing through the system.

Bad:

```vb
Dim value As Variant
```

Good:

```vb
Dim value As String
value = CStr(ws.Cells(1, 1).Value)
```

---

# Object Types

## Worksheet

```vb
Dim ws As Worksheet
```

Use explicit worksheet references. Never rely on `ActiveSheet`.

---

## Workbook

```vb
Dim wb As Workbook
```

Avoid `ActiveWorkbook` whenever possible.

---

## Range

```vb
Dim rng As Range
Dim rngData As Range
```

Always qualify ranges with worksheet context.

---

## Workbook / Worksheet hierarchy rule

Always prefer explicit hierarchy:

```vb
ws.Cells(1, 1)
```

instead of:

```vb
Cells(1, 1)
```

---

# Collections

## Collection

```vb
Dim items As Collection
```

Use for simple ordered lists.

---

## Dictionary (preferred for key-value)

```vb
Dim dictCustomers As Object
Set dictCustomers = CreateObject("Scripting.Dictionary")
```

Use for:

- lookup tables
- caching
- deduplication

---

# Arrays

Use arrays when performance matters.

```vb
Dim values() As String
Dim data() As Variant
```

Prefer arrays over repeated worksheet access.

---

# User Defined Types (UDT)

Use for lightweight structured data.

```vb
Public Type CustomerInfo
    Name As String
    Id As String
    IsActive As Boolean
End Type
```

Use when a class is not necessary.

---

# Classes

Use classes for:

- behavior + data
- encapsulation
- reusable logic
- complex structures

Example:

```vb
Dim customer As Customer
```

---

# Enum Types

Always use enums instead of magic numbers or strings.

```vb
Public Enum LabelType
    Box
    Pallet
    Carton
End Enum
```

Usage:

```vb
Dim labelType As LabelType
labelType = LabelType.Box
```

---

# Constants

Use constants instead of hardcoded values.

```vb
Private Const MAX_RETRIES As Long = 3
Private Const DEFAULT_SHEET As String = "Config"
```

---

# Type Conversion Rules

Always convert explicitly.

## String conversion

```vb
CStr(value)
```

## Numeric conversion

```vb
CLng(value)
CDbl(value)
```

## Boolean conversion

```vb
CBool(value)
```

## Date conversion

```vb
CDate(value)
```

Never rely on implicit conversion.

Bad:

```vb
Dim total As Double
total = ws.Cells(1, 1).Value
```

Good:

```vb
Dim total As Double
total = CDbl(ws.Cells(1, 1).Value)
```

---

# Variant from Excel

Excel `.Value` always returns Variant.

Rule:

> The moment data enters VBA, it must be typed.

Example:

```vb
Dim articleCode As String
articleCode = CStr(ws.Cells(rowIndex, 1).Value)
```

---

# Null, Empty, and Nothing

## Nothing (objects)

```vb
If ws Is Nothing Then Exit Sub
```

## Empty (uninitialized variant)

```vb
If IsEmpty(value) Then
```

## Null (database / COM)

```vb
If IsNull(value) Then
```

Never mix these concepts.

---

# Performance Considerations

- `Long` is faster than `Variant`
- typed variables reduce runtime conversion
- arrays are faster than repeated Range access
- `Dictionary` is faster than looping collections for lookup

---

# Memory Considerations

Avoid unnecessary `Variant` usage:

- increases memory footprint
- reduces performance
- hides type errors

---

# Excel-Specific Rule

Always assume Excel returns `Variant`.

Therefore:

> Every read from a worksheet must be cast.

Example:

```vb
Dim quantity As Long
quantity = CLng(ws.Cells(rowIndex, 2).Value)
```

---

# Defensive Typing

Always validate before casting:

```vb
If IsNumeric(ws.Cells(rowIndex, 2).Value) Then
    quantity = CLng(ws.Cells(rowIndex, 2).Value)
End If
```

---

# AI Rules

When generating VBA code, AI must:

- never default to `Variant`
- always choose the most specific type
- use `Long` instead of `Integer`
- explicitly convert Excel inputs
- avoid implicit type conversion
- prefer `Dictionary` for key-value structures
- use arrays for performance-critical operations
- use enums instead of magic values
- avoid untyped object variables when possible
- ensure every variable has a clear and justified type

---

# Example

Poor:

```vb
Dim x As Variant
x = Cells(1, 1).Value
```

Good:

```vb
Dim articleCode As String
articleCode = CStr(ws.Cells(1, 1).Value)
```

---

# Golden Rules

1. Never use `Variant` by default.
2. Always prefer explicit types.
3. Use `Long` instead of `Integer`.
4. Convert Excel values explicitly.
5. Use `Currency` for financial precision.
6. Use `Dictionary` for key-value logic.
7. Use enums instead of magic values.
8. Validate before converting.
9. Keep types consistent across modules.
10. Strong typing makes VBA safer and faster.