# 04 - Naming Conventions
---
## Objective

Consistent naming is one of the biggest contributors to maintainable VBA code.

A name should describe **what something represents**, not **how it is implemented**.

Good names reduce comments, reduce bugs and make code self-documenting.

---

# General Principles

Always prefer:

- descriptive names
- complete words
- consistency
- readability

Avoid:

- cryptic abbreviations
- Hungarian notation everywhere
- meaningless suffixes
- generic names like:

```
tmp
data
var
obj
thing
test
value
```

---

# Variables

## Local variables

Use camelCase.

Good:

```vb
customerName
currentRow
lastColumn
filePath
totalWeight
```

Bad:

```vb
a
x
tmp
v1
myVariable123
```

---

## Boolean variables

Always read like a question.

Good:

```vb
isValid
hasError
canExport
shouldPrint
isVisible
isLoaded
```

Avoid:

```vb
flag
ok
result
status
```

---

## Collections

Use plural names.

```vb
orders
customers
labels
files
rows
```

Single item:

```vb
customer
order
file
```

---

# Constants

Constants use PascalCase.

```vb
MaxRetries
DefaultTimeout
LabelWidth
```

If module-private:

```vb
Private Const DefaultFontSize As Long = 10
```

---

# Procedures

## Subs

Use a verb.

```vb
ExportOrders
PrintLabel
RefreshData
CreateWorkbook
UpdateStatus
```

Never:

```vb
Data
Button1
Process
Run1
```

---

## Functions

Function names describe the returned value.

Good:

```vb
GetCustomerName
BuildZplLabel
CalculateWeight
FindLastRow
FormatDate
```

Avoid:

```vb
DoCalculation
Execute
Process
Function1
```

---

# Parameters

Parameter names should be meaningful.

Good:

```vb
rowIndex
worksheet
customerId
outputPath
```

Bad:

```vb
i
j
x
p
```

---

# Loop Variables

Small loops may use:

```vb
i
j
k
```

Nested loops:

```vb
rowIndex
columnIndex
```

Collections:

```vb
For Each file In files

For Each customer In customers
```

Never:

```vb
For Each x In y
```

---

# Worksheet Variables

Use:

```vb
ws
wsSource
wsTarget
wsConfig
wsData
```

Avoid:

```vb
sheet1
worksheet2
```

---

# Workbook Variables

```vb
wb
wbSource
wbTarget
wbTemplate
```

---

# Range Variables

```vb
rng
rngData
rngHeader
rngOutput
```

---

# Dictionary Variables

```vb
dictCustomers
dictArticles
dictCache
```

---

# Collection Variables

```vb
customers
orders
labels
files
```

---

# Object Variables

Prefix with their type only when it improves readability.

```vb
httpRequest
xmlDoc
jsonParser
```

Avoid unnecessary prefixes.

Bad:

```vb
objCustomer
objWorkbook
```

---

# Enum Names

Use PascalCase.

```vb
Public Enum LabelType
    Box
    Pallet
    Carton
End Enum
```

---

# User Defined Types

```vb
Public Type CustomerInfo
```

Not:

```vb
custType
```

---

# Module Names

Modules should describe their responsibility.

Examples:

```
LabelBuilder
ExcelHelpers
WorksheetHelpers
SapImport
FileSystemHelpers
DateHelpers
Configuration
```

Avoid:

```
Module1
Utilities2
Test
```

---

# Class Names

Use singular nouns.

```vb
Customer
LabelPrinter
Configuration
Logger
```

---

# Public Members

Public procedures should have readable API names.

```vb
GenerateLabels
ExportWorkbook
LoadConfiguration
```

Think of public procedures as part of a library.

---

# Private Helpers

Private helper names may be slightly more implementation-oriented.

Examples:

```vb
NormalizeArticleNumber
AppendTextField
WriteBarcode
ReadConfigurationValue
```

---

# Temporary Variables

Keep them short-lived.

```vb
currentValue
lineText
cellValue
```

Avoid keeping "tmp" variables alive for dozens of lines.

---

# Naming Excel Tables

Use meaningful table names.

```
tblOrders
tblCustomers
tblArticles
tblSettings
```

---

# Naming Named Ranges

```
ConfigPath
DefaultPrinter
CurrentUser
```

---

# File Names

Good:

```
LabelBuilder.bas
ExcelHelpers.bas
SapImport.bas
Configuration.bas
```

Avoid:

```
Module3.bas
Code.bas
Misc.bas
```

---

# Acronyms

Keep common acronyms uppercase.

Examples:

```
URL
HTTP
XML
JSON
SAP
ZPL
PDF
CSV
SQL
VBA
```

Examples:

```vb
BuildZplLabel
ExportToPdf
ParseJson
ImportSapOrders
```

---

# Abbreviations

Avoid abbreviations unless universally understood.

Acceptable:

```
cfg
msg
qty
min
max
avg
```

Avoid:

```
art
cust
itm
prc
```

unless they are company standards.

---

# AI Naming Rules

When generating VBA code, AI should:

- prefer explicit names
- avoid meaningless abbreviations
- keep naming consistent throughout the project
- reuse existing naming conventions
- never invent multiple names for the same concept
- keep singular/plural consistent
- avoid one-letter variables except loop indexes

---

# Example

Poor:

```vb
Dim x As Variant
Dim t As Worksheet

For i = 2 To l
    x = t.Cells(i, 1)
Next
```

Good:

```vb
Dim wsOrders As Worksheet
Dim currentArticle As String
Dim rowIndex As Long

For rowIndex = 2 To lastRow
    currentArticle = wsOrders.Cells(rowIndex, 1).Value
Next rowIndex
```

---

# Golden Rules

1. Code should read like English.
2. Names should describe intent.
3. Prefer clarity over brevity.
4. Keep naming consistent.
5. Avoid abbreviations.
6. Boolean names answer a question.
7. Collections are plural.
8. Procedures use verbs.
9. Functions describe returned values.
10. Good naming reduces comments.

