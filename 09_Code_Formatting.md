# 09 - Code Formatting

## Objective

Consistent formatting improves readability, maintainability, and collaboration.

Code should be easy to read without requiring comments to understand its structure.

Formatting rules are mandatory across the entire project.

---

# General Principles

Formatting should emphasize:

- readability
- consistency
- visual hierarchy
- simplicity

Every module should look as though it was written by a single developer.

---

# Indentation

Use **4 spaces** for each indentation level.

Never use tabs.

Correct:

```vb
If isValid Then

    ExportData

End If
```

Incorrect:

```vb
If isValid Then
ExportData
End If
```

---

# Blank Lines

Use blank lines to separate logical sections.

Example:

```vb
Dim lastRow As Long
Dim currentRow As Long

lastRow = FindLastRow(ws)

For currentRow = 2 To lastRow

    ExportRow ws, currentRow

Next currentRow

Utils_Log.Info "Export completed"
```

Avoid multiple consecutive blank lines.

---

# One Statement Per Line

Always write one instruction per line.

Correct:

```vb
Dim ws As Worksheet
Dim lastRow As Long
```

Avoid:

```vb
Dim ws As Worksheet: Dim lastRow As Long
```

---

# Line Length

Aim for a maximum of **100 characters**.

Break long statements using VBA's continuation character.

Correct:

```vb
Utils_Log.Info _
    "Export started for worksheet '" & ws.Name & "'"
```

Avoid extremely long lines requiring horizontal scrolling.

---

# Line Continuation

Break long instructions after commas or operators.

Example:

```vb
Call ExportWorksheet( _
    ws, _
    outputFolder, _
    overwriteExisting)
```

Do not split expressions arbitrarily.

---

# Procedure Spacing

Leave one blank line between declarations and executable code.

Example:

```vb
Public Sub Export()

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = Worksheets("Orders")

    lastRow = FindLastRow(ws)

End Sub
```

---

# Variable Declarations

Declare one variable per line.

Correct:

```vb
Dim ws As Worksheet
Dim rowIndex As Long
Dim article As String
```

Avoid:

```vb
Dim ws As Worksheet, rowIndex As Long, article As String
```

---

# Option Explicit

Every module must begin with:

```vb
Option Explicit
```

No exceptions.

---

# Declaration Order

Use the following order:

```vb
Option Explicit

Private Const ...

Public Enum ...

Private Type ...

Private variables

Public procedures

Private procedures
```

Maintain this order throughout the project.

---

# Control Structures

Always place keywords on separate lines.

Example:

```vb
If hasError Then

    Exit Sub

End If
```

Never write:

```vb
If hasError Then Exit Sub
```

unless it is an extremely simple guard clause.

---

# Select Case

Align all cases.

Example:

```vb
Select Case labelType

    Case "A"

        PrintA

    Case "B"

        PrintB

    Case Else

        RaiseUnknownType

End Select
```

---

# Loops

Keep loops visually simple.

Example:

```vb
For rowIndex = 2 To lastRow

    ProcessRow ws, rowIndex

Next rowIndex
```

Always include the loop variable after `Next`.

---

# With Blocks

Use `With` only when several members of the same object are accessed.

Example:

```vb
With ws

    .Cells(1, 1).Value = "Header"
    .Columns(1).AutoFit

End With
```

Avoid nested `With` blocks.

---

# Parentheses

Use parentheses only when required or when they improve readability.

Example:

```vb
If (quantity > 0) And isAvailable Then
```

Avoid unnecessary nesting.

---

# Operators

Always put spaces around operators.

Correct:

```vb
total = price * quantity
```

Avoid:

```vb
total=price*quantity
```

---

# Commas

Always place one space after commas.

Correct:

```vb
ExportLabel ws, rowIndex, printerName
```

Avoid:

```vb
ExportLabel ws,rowIndex,printerName
```

---

# Alignment

Do not vertically align variable names or equal signs.

Correct:

```vb
Dim ws As Worksheet
Dim rowIndex As Long
Dim article As String
```

Avoid:

```vb
Dim ws       As Worksheet
Dim rowIndex As Long
Dim article  As String
```

Automatic formatting should never modify alignment.

---

# Literal Values

Avoid magic numbers.

Prefer:

```vb
Private Const HEADER_ROW As Long = 1

lastRow = HEADER_ROW
```

instead of:

```vb
lastRow = 1
```

---

# Boolean Expressions

Write positive conditions whenever possible.

Prefer:

```vb
If isValid Then
```

instead of:

```vb
If Not isInvalid Then
```

---

# Guard Clauses

Use guard clauses to reduce nesting.

Preferred:

```vb
If ws Is Nothing Then Exit Sub

If lastRow < 2 Then Exit Sub

ProcessWorksheet ws
```

Avoid deeply nested code.

---

# Regions

Do not simulate code regions using comments.

Avoid:

```vb
'=========================
' Export
'=========================
```

Use small procedures instead.

---

# Trailing Spaces

Never leave trailing spaces.

Keep files clean.

---

# Empty Lines at End of File

End every module with a single blank line.

---

# Consistent Capitalization

Use VBA keywords consistently.

Example:

```vb
If
Then
Else
End If

For
Next

Select Case
```

Do not mix capitalization styles.

---

# AI Formatting Rules

When generating VBA code, AI should:

- use 4-space indentation
- insert blank lines between logical blocks
- declare one variable per line
- keep lines under 100 characters
- use meaningful spacing around operators
- avoid inline statements
- use continuation characters correctly
- always include `Option Explicit`
- follow the standard declaration order
- produce consistently formatted modules

---

# Example

Poor:

```vb
Option Explicit
Dim i As Long,j As Long
For i=1 To 10
If Cells(i,1)<>"" Then Cells(i,2)=Cells(i,1)
Next
```

Good:

```vb
Option Explicit

Public Sub CopyValues()

    Dim rowIndex As Long

    For rowIndex = 1 To 10

        If Cells(rowIndex, 1).Value <> "" Then

            Cells(rowIndex, 2).Value = Cells(rowIndex, 1).Value

        End If

    Next rowIndex

End Sub
```

---

# Golden Rules

1. Formatting must be consistent across the project.
2. Use 4 spaces for indentation.
3. Declare one variable per line.
4. Separate logical blocks with blank lines.
5. One statement per line.
6. Avoid horizontal scrolling.
7. Prefer guard clauses over deep nesting.
8. Avoid decorative formatting.
9. Keep procedures visually simple.
10. Readability always comes before compactness.