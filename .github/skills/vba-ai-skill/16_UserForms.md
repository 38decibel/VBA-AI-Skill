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

> UserForms only collect input and trigger actions. They do not process business logic.

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
- implement UI logic (enable a control when the command can run, show a label when the model is invalid)
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

# 4. Instance Lifecycle

A UserForm module is a class that comes with a global **default instance** carrying the form's name.
Code that refers to that global name acts on the default instance, which is not necessarily the one displayed
(and not the one you get if the form is created with `New`).

Rules:

- **Never** show or use the default instance (`frmExport.Show`).
- Treat the default instance like the default instance of any class: stateless, never displayed.
- Always create a new instance and show that instance:

```vb
Dim dialog As frmExport
Set dialog = New frmExport
dialog.Show
```

- The code that creates the form is responsible for its lifetime, so it can safely read the form's model afterwards.

---

# 5. Hide, Never Self-Destruct (`QueryClose`)

The [X] button of the control box **destroys** the form instance. Code that created the form then holds an
invalid reference, and code that reads the form's controls after closing fails.

Every modal form handles `QueryClose` so that the only way to close it is to **hide** it:

```vb
Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OnFormCancelled
    End If

End Sub

Private Sub OnFormCancelled()

    m_model.IsCancelled = True
    Me.Hide

End Sub
```

Buttons use `Me.Hide` as well; the form is never `Unload`ed by its own code.

---

# 6. Form Model

Extract the form's data into a model class instead of reading and writing controls from outside.
Controls manipulate the model; the caller consumes the model.

A model typically exposes:

- a read/write property for each editable field
- read-only properties for data the controls need (items of a list box)
- an `IsCancelled` flag set when the user cancels
- an `IsValid` property that returns `True` when all required values are present and valid

```vb
Option Explicit

Private m_orderId As String

Public IsCancelled As Boolean ' replace with a property in production code

Public Property Get OrderId() As String
    OrderId = m_orderId
End Property

Public Property Let OrderId(ByVal newValue As String)
    m_orderId = newValue
End Property

Public Property Get IsValid() As Boolean
    IsValid = Len(m_orderId) > 0
End Property
```

(Fields are never public in production classes, chapter 15: expose `IsCancelled` through a `Property Get/Let`.)

`IsValid` drives the enabled state of the Accept button, so invalid input can never be submitted.

---

# 7. Modal vs Modeless

- A **modal** form is a transactional dialog: show it, wait, then consume the model (Accept) or do nothing (Cancel).
  This is the default and the easiest to reason about.
- A **modeless** form returns immediately and interactions become asynchronous (event-driven). Use it only when needed,
  and give it a presenter object that owns and shows the form instance.

---

# 8. Event-Driven Architecture

UserForms are event-driven.

Typical events:

- `UserForm_Initialize`
- `CommandButton_Click`
- `ComboBox_Change`

Each event must remain **small and orchestrated**.

---

# 9. Button Click Pattern (Standard)

A button click is an **entry point**: it needs a handler (chapter 05).

## Correct structure

```vb
Private Sub btnExport_Click()

    On Error GoTo CleanFail

    Utils_Log.Info "frmExport", "Export button clicked"

    ExportController.RunExport m_model

    Me.Hide

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "frmExport.btnExport_Click"
    Resume CleanExit

End Sub
```

A command button should invoke a command, not implement side effects:
the logic lives in a controller module (procedural) or in a command object (`ICommand`, chapter 28).

---

# 10. No Business Logic in UI

Bad:

```vb
If txtQty.Value > 10 Then
    discount = 0.1
End If
```

Good:

```vb
discount = order.CalculateDiscount()
```

---

# 11. UserForm to Controller Pattern

UserForms must delegate logic:

```
UserForm -> Module (Controller) -> Class Modules -> Excel
```

Example:

```vb
ExportController.RunExport m_model
```

---

# 12. The Calling Code (Presenter / Controller)

```vb
Public Sub ShowExportDialog()

    Dim model As clsExportModel
    Set model = New clsExportModel

    Dim dialog As frmExport
    Set dialog = New frmExport
    dialog.Init model

    dialog.Show

    If model.IsCancelled Then Exit Sub

    ExportController.RunExport model

End Sub
```

The form's `Init` method stores the model, configures controls from it, and is where the form is wired.

---

# 13. Data Binding Rules

## Load data into UI from the model

```vb
txtOrderId.Value = m_model.OrderId
```

## Never bind UI directly to Excel ranges

Bad:

```vb
txtOrderId.Value = ws.Cells(1, 1).Value
```

Good:

```vb
txtOrderId.Value = m_model.OrderId
```

Change handlers write to the model:

```vb
Private Sub txtOrderId_Change()

    m_model.OrderId = txtOrderId.Value
    RefreshState

End Sub

Private Sub RefreshState()
    btnOk.Enabled = m_model.IsValid
End Sub
```

---

# 14. Validation Rules

UserForms may validate:

- empty fields
- format correctness
- numeric input

Never trust user input: assume values are empty, mistyped, out of range, in another locale's number format,
or that the dialog is cancelled.

Show validation errors next to the field (an icon or a colored hint with a tooltip), not only in a `MsgBox`.
Use background colors on input controls only to signal something clearly, such as a validation error.

## Forbidden validation

- business rules
- pricing logic
- SAP constraints
- export rules

---

# 15. MessageBox Usage

## Allowed (limited)

- input validation
- user feedback

## Forbidden

- logging
- debugging
- system errors (use Utils_Log instead)

---

# 16. Logging in UserForms

Important actions are logged:

```vb
Utils_Log.Info "frmExport", "User started export"
```

Errors in event handlers:

```vb
Utils_Log.Error Err, "frmExport.btnExport_Click"
```

---

# 17. Initialization Pattern

```vb
Public Sub Init(ByVal model As clsExportModel)

    Utils_Guard.NotNothing model, "model"

    Set m_model = model

    txtOrderId.Value = m_model.OrderId
    RefreshState

End Sub
```

Do not rely on `UserForm_Initialize` for model-dependent setup: the model does not exist yet at that point.
Do not rely on `UserForm_Activate` either: it fires again every time the user returns to the form.

---

# 18. Avoid Heavy Processing in Forms

Bad:

```vb
For Each row In lo.ListRows
    ' processing logic
Next row
```

Good:

```vb
ExportController.RunExport m_model
```

---

# 19. UI State Management

UserForms may store temporary UI state:

```vb
Private m_selectedFilePath As String
```

Rules:

- state must be UI-only
- never store business data (it lives in the model)
- fields are private, prefixed `m_`

---

# 20. Naming Forms and Controls

Name **everything** the code or a future maintainer may interact with.
Default names (`UserForm1`, `CommandButton1`, `Label42`, `Rounded Rectangle 1`) are forbidden.

For UserForms, the project accepts a short type prefix on forms and controls, because controls are
referenced constantly and share one flat namespace. This is the only place where a type prefix is used.

| Element | Prefix | Example |
|---------|--------|---------|
| UserForm | `frm` | `frmExport` |
| Button | `btn` | `btnExport`, `btnCancel` |
| Text box | `txt` | `txtOrderId` |
| Combo box | `cmb` | `cmbPrinter` |
| List box | `lst` | `lstOrders` |
| Check box | `chk` | `chkOverwrite` |
| Option button | `opt` | `optMetric` |
| Label | `lbl` | `lblInstructions` |
| Frame | `fra` | `fraOptions` |

Event handlers then read naturally: `btnExport_Click`, `txtOrderId_Change`.

Shapes on worksheets that run a macro get a purposeful name as well (`ExportButton`), never the default one.

Set a logical `TabIndex` on every control so the form can be navigated with the keyboard,
and provide a tooltip on each field the user must fill in (state whether it is required and which format it expects).

---

# 21. Reusability Rule

UserForms must be reusable:

- no hardcoded worksheet names
- no hardcoded business rules
- no embedded logic dependencies

---

# 22. Recommended Architecture Flow

```
UserForm
   |
Controller Module
   |
Class Modules (Business Logic)
   |
Excel Object Model
   |
Utils_Log (traceability)
```

---

# 23. Performance Rules

- avoid loops in UI
- avoid Excel access in forms
- keep event handlers lightweight
- delegate all processing

---

# 24. AI Rules

When generating UserForms, AI must:

- never include business logic in forms
- never show or use the default instance: create a new instance and show that instance
- always handle `QueryClose` and hide the form instead of letting it self-destruct
- use a model class (with `IsValid` and `IsCancelled`) for the form's data
- always delegate to controller modules
- keep event handlers minimal and put a handler in each entry-point event
- use Utils_Log for all logging
- avoid Excel direct manipulation
- validate only UI-level input
- never embed loops over datasets
- ensure strict UI/business separation
- avoid MsgBox for system errors
- name every form and control with the project prefixes
- maintain event-driven structure

---

# Example

```vb
Private Sub btnGenerate_Click()

    On Error GoTo CleanFail

    Utils_Log.Info "frmExport", "Generate clicked"

    ExportController.RunExport m_model

    Me.Hide

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "frmExport.btnGenerate_Click"
    Resume CleanExit

End Sub
```

---

# Golden Rules

1. UserForms are UI only.
2. No business logic in forms.
3. Always create a new instance; never show the default instance.
4. Hide the form; never let it self-destruct (`QueryClose`).
5. Form data lives in a model class with `IsValid` and `IsCancelled`.
6. Always delegate to controllers.
7. Keep event handlers small.
8. Validate only user input, and never trust it.
9. Name every form and control.
10. UI must remain thin and replaceable.
