# 04 - Naming Conventions

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
- English identifiers across modules, procedures, and variables

Avoid:

- cryptic abbreviations
- Hungarian notation everywhere (type prefixes such as `str`, `lng`, `obj`; the only accepted type-like
  prefixes are `cls` on classes and `frm`, `btn`, `txt`... on UserForms and controls, chapter 16)
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

Use camelCase by default.

Alternative styles (lowercase_with_underscores / UPPERCASE_WITH_UNDERSCORES)
are allowed when required by external interfaces or existing module style.

## Underscores

An underscore in a procedure name is reserved for **event handlers** (`btnExport_Click`, `Worksheet_Change`)
and for the module-domain prefixes of this project (`Utils_Log`, `SAP_Session`). Anywhere else, a name with an
underscore looks like an event handler or an interface member and confuses both the reader and the VBE
(an `Implements` member is named `IShape_Draw`). Use PascalCase for procedures and camelCase for variables.

## One identifier, one purpose

A variable has one meaning for its whole life. Never reuse `i`, `temp` or `result` for a second purpose
in the same procedure: declare a new, well-named variable instead.

## Scope prefixes

| Scope | Prefix | Example |
|-------|--------|---------|
| Module-level (private) | `m_` | `m_orderId` |
| Global (`Public` in a standard module, avoid) | `g_` | `g_Config` |
| Constant | `C_` or UPPER_CASE (project style) | `C_DEFAULT_TIMEOUT` |
| `ByRef` output parameter | `out` | `outWorksheet` |

The prefix `p` is not used for class fields.

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

Constants use PascalCase by default.

UPPERCASE_WITH_UNDERSCORES is allowed for shared symbolic constants.

For module-private constants, prefix with `C_`.

```vb
MaxRetries
DefaultTimeout
LabelWidth
```

If module-private:

```vb
Private Const C_DefaultFontSize As Long = 10
```

---

# Procedures

## Subs

Use a verb.

PascalCase is the default in this repository.
lowerCamelCase is acceptable when preserving legacy module consistency.

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

Standard modules use a domain prefix followed by a responsibility (see chapter 03):

```
Utils_Log
Utils_Guard
Excel_Table
SAP_Session
Business_LabelBuilder
```

A macro entry point module is named after the macro (`ExportLabels`), one macro per module (chapter 06).

Avoid:

```
Module1
Utilities2
Test
```

---

# Class Names

Class modules use the `cls` prefix and a singular noun:

```vb
clsCustomer
clsLabelPrinter
clsConfiguration
clsLogger
```

Interfaces (abstract classes used with `Implements`) use the `I` prefix without `cls`:

```vb
ICommand
IFileProvider
ILogger
```

See chapter 15 and chapter 28.

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
- use `m_` for module-level fields, `out` for output parameters, `cls` for classes, `I` for interfaces
- never use underscores in procedure names except for event handlers
- never reuse a variable for a second purpose

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
    currentArticle = wsOrders.Cells(rowIndex, 1).Value2
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


