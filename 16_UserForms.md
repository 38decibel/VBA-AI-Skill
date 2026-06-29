# 16 - UserForms

## Objective

UserForms are the **UI layer of VBA applications**.

They are responsible for:

- user interaction
- input validation (basic)
- triggering business workflows
- displaying structured data

UserForms must NEVER contain business logic.

They are strictly a **presentation layer**.

---

# Golden Rule

> UserForms only collect input and trigger actions — they do not process business logic.

---

# 1. Layer Separation Principle

VBA architecture must be separated into:

| Layer | Responsibility |
|------|----------------|
| UserForm | UI / interaction |
| Class Modules | business logic |
| Modules | orchestration |
| Utils_Log | logging |

---

# 2. UserForm Responsibilities

A UserForm can:

- display data
- collect user input
- validate basic fields (format only)
- call business procedures
- show messages (limited)

---

## Example responsibilities

- selecting a file
- entering an order ID
- choosing an export mode
- triggering "Generate Labels"

---

# 3. What UserForms MUST NOT do

UserForms must NOT:

- contain business rules
- manipulate Excel data directly (except UI binding)
- contain loops over datasets
- perform calculations
- implement export logic
- access ListObjects directly for transformation

Bad:

```vb
For i = 1 To 1000
    ws.Cells(i, 1).Value = "X"
Next i
```

---

# 4. Event-Driven Architecture

UserForms are event-driven.

Typical events:

- `UserForm_Initialize`
- `CommandButton_Click`
- `ComboBox_Change`

Each event must remain **small and orchestrated**.

---

# 5. Button Click Pattern (Standard)

## Correct structure

```vb
Private Sub btnExport_Click()

    On Error GoTo ErrorHandler

    Utils_Log.Info "UserForm", "Export button clicked"

    Call ExportController.RunExport(txtOrderId.Value)

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "UserForm.btnExport_Click"
End Sub
```

---

# 6. No Business Logic in UI

Bad:

```vb
If txtQty.Value > 10 Then
    discount = 0.1
End If
```

Good:

```vb
discount = order.CalculateDiscount(txtQty.Value)
```

---

# 7. UserForm → Controller Pattern

UserForms must delegate logic:

```
UserForm → Module (Controller) → Class Modules → Excel
```

Example:

```vb
Call ExportController.RunExport(orderId)
```

---

# 8. Data Binding Rules

## Load data into UI

```vb
txtOrderId.Value = order.Id
txtQty.Value = order.Quantity
```

## Never bind UI directly to Excel ranges

Bad:

```vb
txtOrderId.Value = ws.Cells(1, 1).Value
```

Good:

```vb
txtOrderId.Value = order.GetId()
```

---

# 9. Validation Rules

UserForms may validate:

- empty fields
- format correctness
- numeric input

## Allowed validation

```vb
If txtOrderId.Value = "" Then
    MsgBox "Order ID is required"
    Exit Sub
End If
```

---

## Forbidden validation

- business rules
- pricing logic
- SAP constraints
- export rules

---

# 10. MessageBox Usage

## Allowed (limited)

- input validation
- user feedback

## Forbidden

- logging
- debugging
- system errors (use Utils_Log instead)

---

# 11. Logging in UserForms

All actions must be logged:

```vb
Utils_Log.Info "UserForm", "User started export"
```

Errors:

```vb
Utils_Log.Error Err, "UserForm.btnExport_Click"
```

---

# 12. Initialization Pattern

```vb
Private Sub UserForm_Initialize()

    Utils_Log.Debug "UserForm", "Initializing form"

    LoadDefaultValues

End Sub
```

---

# 13. Avoid Heavy Processing in Forms

Bad:

```vb
For Each row In lo.ListRows
    ' processing logic
Next row
```

Good:

```vb
Call ExportController.RunExport
```

---

# 14. UI State Management

UserForms may store temporary UI state:

```vb
Private selectedFilePath As String
```

Rules:

- state must be UI-only
- never store business data
- reset on form close if needed

---

# 15. Reusability Rule

UserForms must be reusable:

- no hardcoded worksheet names
- no hardcoded business rules
- no embedded logic dependencies

---

# 16. Separation of Concerns Example

## Bad

UserForm does everything:

- reads Excel
- processes data
- exports file

---

## Good

UserForm only:

- collects input
- calls controller
- displays result

---

# 17. Recommended Architecture Flow

```
UserForm
   ↓
Controller Module
   ↓
Class Modules (Business Logic)
   ↓
Excel Object Model
   ↓
Utils_Log (traceability)
```

---

# 18. Performance Rules

- avoid loops in UI
- avoid Excel access in forms
- keep event handlers lightweight
- delegate all processing

---

# 19. Error Handling Pattern

Every event must include error handling:

```vb
On Error GoTo ErrorHandler

' logic

Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "UserForm.Event"
End Sub
```

---

# 20. AI Rules

When generating UserForms, AI must:

- never include business logic in forms
- always delegate to controller modules
- keep event handlers minimal
- use Utils_Log for all logging
- avoid Excel direct manipulation
- validate only UI-level input
- never embed loops over datasets
- ensure strict UI/business separation
- avoid MsgBox for system errors
- maintain event-driven structure

---

# Example

```vb
Private Sub btnGenerate_Click()

    On Error GoTo ErrorHandler

    Utils_Log.Info "UserForm", "Generate clicked"

    Call ExportController.RunExport(txtOrderId.Value)

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "UserForm.btnGenerate_Click"
End Sub
```

---

# Golden Rules

1. UserForms are UI only.
2. No business logic in forms.
3. Always delegate to controllers.
4. Keep event handlers small.
5. Validate only user input.
6. Use logging for traceability.
7. Avoid Excel manipulation in UI.
8. Never process datasets in forms.
9. Use MsgBox only for user feedback.
10. UI must remain thin and replaceable.