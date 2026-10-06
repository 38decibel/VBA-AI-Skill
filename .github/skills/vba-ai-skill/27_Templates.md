# 27 - Templates (VBA Code Patterns Library)

## Objective

This chapter provides **standard reusable templates** for the VBA AI Skill framework.

These templates ensure:

- consistency
- performance
- safety
- architecture compliance
- fast development

They are the **reference implementations** that all generated code must follow.

---

# Golden Rule

> Never reinvent patterns already defined in templates.

Conventions shared by all templates:

- labels are always `CleanExit` and `CleanFail`
- no `Call` keyword, and no parentheses around the arguments of a `Sub` call
- parameters are always explicitly `ByVal` (or `ByRef` when intentional)
- variables are declared where they are first needed, one `Dim` per variable
- contract violations raise through `Utils_Guard`; expected data situations are logged as `Warning`
- `Utils_Log` signatures follow chapter 10: `Info source, message` and `Error Err, source`

---

# 1. Entry Point Template (macro, button, `Workbook_Open`)

Handler mandatory: log once, resume to the single exit point.

```vb
Public Sub ExportLabels()

    On Error GoTo CleanFail

    Utils_Log.Info "ExportLabels", "Start"

    ' Main logic goes here.

    Utils_Log.Info "ExportLabels", "Completed"

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "ExportLabels"
    Resume CleanExit

End Sub
```

For a user-initiated entry point, add the user message in `CleanFail`:

```vb
CleanFail:
    Utils_Log.Error Err, "ExportLabels"
    MsgBox _
        "The operation could not be completed." & vbCrLf & _
        "See the application log for details.", _
        vbExclamation
    Resume CleanExit
```

---

# 2. Procedure That Owns State (clean up, then re-raise)

Use when the procedure changes Excel settings, holds a COM object, a SAP session or a file handle.
The error is **not** logged here: the caller that finally handles it logs it once.

```vb
Public Sub LoadOrders(ByVal ws As Worksheet)

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Utils_Guard.NotNothing ws, "ws"

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Main logic goes here.

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If

    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit

End Sub
```

---

# 3. Plain Helper Function (no handler)

No resource, no recovery: validate and let errors propagate.

```vb
Public Function NormalizeArticle(ByVal article As String) As String

    Utils_Guard.NotEmpty article, "article"

    NormalizeArticle = UCase$(Trim$(article))

End Function
```

---

# 4. Try Function (expected failure)

`Resume Next` lives in a tiny dedicated function; the result is returned through an `out` parameter.

```vb
Public Function TryGetWorksheet( _
    ByVal wb As Workbook, _
    ByVal sheetName As String, _
    ByRef outWorksheet As Worksheet) As Boolean

    Utils_Guard.NotNothing wb, "wb"

    Set outWorksheet = Nothing

    On Error Resume Next
    Set outWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    TryGetWorksheet = Not outWorksheet Is Nothing

End Function
```

Call site:

```vb
Dim wsOrders As Worksheet

If Not Utils_Excel.TryGetWorksheet(wb, "Orders", outWorksheet:=wsOrders) Then
    Utils_Log.Warning "ImportOrders", "Orders worksheet not found"
    Exit Sub
End If
```

---

# 5. Excel Bulk Processing Template (Value2)

Plain helper: guards and data situations, no handler. The table is assumed to have at least two columns,
so `Value2` is always a two-dimensional array.

```vb
Public Sub ProcessTable(ByVal lo As ListObject)

    Utils_Guard.NotNothing lo, "lo"

    If lo.DataBodyRange Is Nothing Then

        Utils_Log.Warning "ProcessTable", "Table has no data rows", "Table=" & lo.Name
        Exit Sub

    End If

    Dim data As Variant
    data = lo.DataBodyRange.Value2

    Dim rowIndex As Long

    For rowIndex = 1 To UBound(data, 1)

        If Len(CStr(data(rowIndex, 1))) > 0 Then
            data(rowIndex, 2) = data(rowIndex, 1)
        End If

    Next rowIndex

    lo.DataBodyRange.Value2 = data

    Utils_Log.Info "ProcessTable", "Completed"

End Sub
```

---

# 6. Dictionary Pattern Template

```vb
Public Function BuildCustomerDictionary(ByVal lo As ListObject) As Object

    Utils_Guard.NotNothing lo, "lo"

    Dim dictCustomers As Object
    Set dictCustomers = CreateObject("Scripting.Dictionary")

    If lo.DataBodyRange Is Nothing Then
        Set BuildCustomerDictionary = dictCustomers
        Exit Function
    End If

    Dim data As Variant
    data = lo.DataBodyRange.Value2

    Dim rowIndex As Long
    Dim customerKey As String

    For rowIndex = 1 To UBound(data, 1)

        customerKey = CStr(data(rowIndex, 1))

        If Not dictCustomers.Exists(customerKey) Then
            dictCustomers.Add customerKey, data(rowIndex, 2)
        End If

    Next rowIndex

    Utils_Log.Debug "BuildCustomerDictionary", "Keys=" & dictCustomers.Count

    Set BuildCustomerDictionary = dictCustomers

End Function
```

---

# 7. UserForm Event Template

Entry point: thin handler that delegates to a controller (chapter 16).

```vb
Private Sub btnExecute_Click()

    On Error GoTo CleanFail

    Utils_Log.Info "UserForm", "Execute clicked"

    ExecuteController.RunExecute txtInput.Value

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "UserForm.btnExecute_Click"
    Resume CleanExit

End Sub
```

Standard form closing (hide, never unload):

```vb
Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OnFormCancelled
    End If

End Sub
```

---

# 8. Event Handler Template (Worksheet)

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo CleanFail

    If Target.CountLarge > 1 Then Exit Sub

    Utils_Log.Debug "Event", "Change detected: " & Target.Address

    EventController.HandleChange Me, Target

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Worksheet_Change"
    Resume CleanExit

End Sub
```

---

# 9. COM Automation Template (Outlook Example)

The COM object is released in `CleanExit`, with or without error. The error is re-raised after cleanup.

```vb
Public Sub SendEmail(ByVal toAddress As String, ByVal subject As String, ByVal body As String)

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Dim outlookApp As Object
    Dim mailItem As Object

    Utils_Guard.NotEmpty toAddress, "toAddress"

    On Error GoTo CleanFail

    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)

    mailItem.To = toAddress
    mailItem.Subject = subject
    mailItem.Body = body
    mailItem.Send

    Utils_Log.Info "COM.Outlook", "Email sent to " & toAddress

CleanExit:
    Set mailItem = Nothing
    Set outlookApp = Nothing

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If

    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit

End Sub
```

---

# 10. File System Template

The file handle is closed in `CleanExit`, so an error never leaves the file locked.

```vb
Public Sub WriteFile(ByVal filePath As String, ByVal content As String)

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Dim fso As Object
    Dim textFile As Object

    Utils_Guard.NotEmpty filePath, "filePath"

    On Error GoTo CleanFail

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textFile = fso.OpenTextFile(filePath, 8, True)

    textFile.WriteLine content

    Utils_Log.Info "FileSystem", "File written: " & filePath

CleanExit:
    If Not textFile Is Nothing Then textFile.Close

    Set textFile = Nothing
    Set fso = Nothing

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If

    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit

End Sub
```

---

# 11. Guard Clause Template

Contract violations raise a custom error. Expected data situations log a warning and exit.

```vb
Public Sub ExportWorksheet(ByVal ws As Worksheet, ByVal outputFolder As String)

    Utils_Guard.NotNothing ws, "ws"
    Utils_Guard.NotEmpty outputFolder, "outputFolder"

    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row < 2 Then

        Utils_Log.Warning "ExportWorksheet", "No data rows", "Sheet=" & ws.Name
        Exit Sub

    End If

    ' logic here

End Sub
```

The `Utils_Guard` module and the `AppError` enumeration are defined in chapter 05.

---

# 12. Initialization Template

Entry point of the application (composition root, chapter 28).

```vb
Public Sub Initialize()

    On Error GoTo CleanFail

    Utils_Log.Info "Initialize", "Starting initialization"

    LoadConfiguration
    PrepareEnvironment

    Utils_Log.Info "Initialize", "Completed"

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "Initialize"
    Resume CleanExit

End Sub
```

---

# 13. Safe Excel State Template

```vb
Public Sub RunWithSafeExcelState()

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' logic here

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If

    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit

End Sub
```

---

# 14. Orchestration Template (high level, reads like a story)

The public entry point sits at the top and calls private procedures of decreasing abstraction.

```vb
Public Sub MainProcess()

    On Error GoTo CleanFail

    ValidateInput
    ProcessData
    ExportData

    Utils_Log.Info "MainProcess", "Completed"

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "MainProcess"
    Resume CleanExit

End Sub

Private Sub ValidateInput()
    ' ...
End Sub

Private Sub ProcessData()
    ' ...
End Sub

Private Sub ExportData()
    ' ...
End Sub
```

---

# 15. Class Module Template

Private fields use the `m_` prefix. No public fields. No `ActiveSheet`/`Selection`.

```vb
Option Explicit

Private m_value As String

Public Sub Init(ByVal value As String)

    Utils_Guard.NotEmpty value, "value"

    m_value = value

End Sub

Public Property Get Value() As String
    Value = m_value
End Property
```

For classes that must be created already valid (factory method, interfaces, dependency injection),
see chapter 28.

---

# 16. Logging Pattern Template

```vb
Utils_Log.Debug "Module", "Message"
Utils_Log.Info "Module", "Message"
Utils_Log.Warning "Module", "Message", "Context=..."
Utils_Log.Error Err, "Module.Procedure"
```

---

# 17. AI Usage Rule

AI must:

- prefer templates over custom patterns
- not reinvent standard structures
- ensure consistency across modules
- choose the template that matches the procedure role (entry point, state owner, plain helper, `Try`)
- always respect performance templates

---

# 18. Anti-Template Warning

Never generate:

- unstructured entry points without a handler
- handlers that only log and re-raise in plain helpers
- cleanup placed before `Exit Sub` outside `CleanExit`
- the `Call` keyword
- direct Excel cell loops
- COM without cleanup
- unmanaged events

---

# Golden Rules

1. Always use templates as default.
2. Match the template to the procedure role.
3. Never skip cleanup: it belongs in `CleanExit`.
4. Always log handled failures once.
5. Prefer bulk operations.
6. Keep UI, logic, and data separated.
7. Always clean COM objects.
8. Always protect Excel state.
9. Use guard clauses that raise for contract violations.
10. Standardization beats creativity in VBA.
