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

---

# 1. Standard Procedure Template

## Base structure (MANDATORY)

```vb
Public Sub ProcedureName()

    On Error GoTo ErrorHandler

    ' =========================
    ' LOG START
    ' =========================
    Utils_Log.Debug "ProcedureName", "Start"

    ' =========================
    ' LOGIC
    ' =========================

    ' TODO: implementation

    ' =========================
    ' LOG END
    ' =========================
    Utils_Log.Info "ProcedureName", "Completed"

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "ProcedureName"
End Sub
```

---

# 2. Function Template

```vb
Public Function FunctionName(ByVal inputValue As String) As String

    On Error GoTo ErrorHandler

    Dim result As String

    result = ""

    ' logic here

    FunctionName = result
    Exit Function

ErrorHandler:
    Utils_Log.Error Err, "FunctionName"
    FunctionName = ""
End Function
```

---

# 3. Excel Bulk Processing Template (VALUE2)

## Standard pattern

```vb
Public Sub ProcessTable(ByVal lo As ListObject)

    On Error GoTo ErrorHandler

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim data As Variant
    data = lo.DataBodyRange.Value2

    Dim i As Long
    For i = 1 To UBound(data, 1)

        If data(i, 1) <> "" Then
            data(i, 2) = data(i, 1)
        End If

    Next i

    lo.DataBodyRange.Value2 = data

    Utils_Log.Info "ProcessTable", "Completed"

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "ProcessTable"
End Sub
```

---

# 4. Dictionary Pattern Template

```vb
Public Sub BuildDictionary()

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim data As Variant
    data = lo.DataBodyRange.Value2

    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Not dict.Exists(data(i, 1)) Then
            dict.Add data(i, 1), data(i, 2)
        End If
    Next i

    Utils_Log.Info "BuildDictionary", "Keys=" & dict.Count

End Sub
```

---

# 5. UserForm Event Template

```vb
Private Sub btnExecute_Click()

    On Error GoTo ErrorHandler

    Utils_Log.Info "UserForm", "Execute clicked"

    Call Controller.ExecuteAction(txtInput.Value)

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "UserForm.btnExecute_Click"
End Sub
```

---

# 6. Event Handler Template (Worksheet)

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ErrorHandler

    If Target.CountLarge > 1 Then Exit Sub

    Utils_Log.Debug "Event", "Change detected: " & Target.Address

    Call EventController.HandleChange(Me, Target)

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "Worksheet_Change"
End Sub
```

---

# 7. COM Automation Template (Outlook Example)

```vb
Public Sub SendEmail(ByVal toAddr As String, ByVal subject As String, ByVal body As String)

    On Error GoTo ErrorHandler

    Dim olApp As Object
    Dim mail As Object

    Set olApp = CreateObject("Outlook.Application")
    Set mail = olApp.CreateItem(0)

    mail.To = toAddr
    mail.Subject = subject
    mail.Body = body
    mail.Send

    Set mail = Nothing
    Set olApp = Nothing

    Utils_Log.Info "COM.Outlook", "Email sent to " & toAddr

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "COM.Outlook.SendEmail"
End Sub
```

---

# 8. File System Template

```vb
Public Sub WriteFile(ByVal filePath As String, ByVal content As String)

    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 8, True)

    ts.WriteLine content
    ts.Close

    Set ts = Nothing
    Set fso = Nothing

    Utils_Log.Info "FileSystem", "File written: " & filePath

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "FileSystem.WriteFile"
End Sub
```

---

# 9. Safe Guard Clause Template

```vb
Public Sub SafeProcess(ByVal value As Variant)

    If IsEmpty(value) Then Exit Sub
    If value = "" Then Exit Sub
    If Not IsNumeric(value) Then Exit Sub

    ' logic here

End Sub
```

---

# 10. Initialization Template

```vb
Public Sub Init()

    On Error GoTo ErrorHandler

    Utils_Log.Info "Init", "Starting initialization"

    Call LoadConfiguration
    Call PrepareEnvironment

    Utils_Log.Info "Init", "Completed"

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "Init"
End Sub
```

---

# 11. Safe Excel State Template

```vb
Public Sub SafeExcelOperation()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' logic here

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "SafeExcelOperation"
    Resume CleanExit
End Sub
```

---

# 12. Refactoring Template Pattern

```vb
Public Sub MainProcess()

    On Error GoTo ErrorHandler

    Call ValidateInput
    Call ProcessData
    Call ExportData

    Utils_Log.Info "MainProcess", "Completed"

    Exit Sub

ErrorHandler:
    Utils_Log.Error Err, "MainProcess"
End Sub
```

---

# 13. Class Module Template

```vb
Option Explicit

Private pValue As String

Public Sub Init(ByVal value As String)
    pValue = value
End Sub

Public Function GetValue() As String
    GetValue = pValue
End Function
```

---

# 14. Logging Pattern Template

```vb
Utils_Log.Debug "Module", "Message"
Utils_Log.Info "Module", "Message"
Utils_Log.Error Err, "Module.Procedure"
```

---

# 15. AI Usage Rule

AI must:

- prefer templates over custom patterns
- not reinvent standard structures
- ensure consistency across modules
- always include error handling template
- always include logging template
- always respect performance templates

---

# 16. Anti-Template Warning

Never generate:

- unstructured procedures
- missing error handling
- missing logging
- direct Excel cell loops
- COM without cleanup
- unmanaged events

---

# Golden Rules

1. Always use templates as default.
2. Never omit error handling.
3. Always include logging.
4. Prefer bulk operations.
5. Keep UI, logic, and data separated.
6. Always clean COM objects.
7. Always protect Excel state.
8. Use guard clauses.
9. Keep procedures small and focused.
10. Standardization beats creativity in VBA.