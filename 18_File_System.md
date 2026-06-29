# 18 - File System (VBA File Handling)

## Objective

File system operations are a critical part of VBA automation:

- imports / exports
- logs persistence
- report generation
- data ingestion (CSV, TXT, XLSX)
- external integrations

This chapter defines strict rules for **safe, predictable, and portable file handling**.

---

# Golden Rule

> Never hardcode file paths or rely on implicit file context. Always use explicit, validated paths.

---

# 1. Always Use `FileSystemObject` (Preferred)

## Recommended approach

```vb
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
```

Why:
- more flexible than `Dir`
- object-oriented
- better readability
- supports full file operations

---

# 2. Path Management Rule

## NEVER hardcode full paths

Bad:

```vb
"C:\Users\Daniel\Documents\data.xlsx"
```

Good:

```vb
basePath & "\data.xlsx"
```

---

## Recommended pattern

```vb
Dim basePath As String
basePath = ThisWorkbook.Path
```

---

# 3. Use Centralized Path Variables

Paths must be managed in:

- Config sheet (`wsConfig`)
- or configuration class

Example:

```vb
pBaseExportPath
pLogPath
pImportPath
```

---

# 4. File Existence Check

Always validate before access:

```vb
If Not fso.FileExists(filePath) Then
    Utils_Log.Error Err, "File not found: " & filePath
    Exit Sub
End If
```

---

# 5. Folder Existence Check

```vb
If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder folderPath
End If
```

---

# 6. File Opening Rules

## Excel files

```vb
Dim wb As Workbook
Set wb = Workbooks.Open(filePath)
```

## Text files (advanced)

```vb
Dim ts As Object
Set ts = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
```

---

# 7. File Modes (Important)

| Mode | Value | Description |
|------|------|-------------|
| ForReading | 1 | read-only |
| ForWriting | 2 | overwrite |
| ForAppending | 8 | append |

---

# 8. Safe File Write Pattern

```vb
Dim ts As Object
Set ts = fso.OpenTextFile(filePath, 8, True)

ts.WriteLine "log entry"
ts.Close
```

Always close file handles.

---

# 9. File Naming Rules

## Must be:

- deterministic
- timestamped if logs
- descriptive
- system-safe

---

## Example

```vb
"Export_Orders_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
```

---

# 10. Forbidden Characters Handling

Windows invalid filename characters:

```
\ / : * ? " < > |
```

Always sanitize:

```vb
fileName = Replace(fileName, ":", "_")
```

---

# 11. Temporary Files Rule

Use:

```vb
Environ("TEMP")
```

Example:

```vb
tempPath = Environ("TEMP") & "\export.tmp"
```

---

# 12. File Copy Operations

```vb
fso.CopyFile sourcePath, destinationPath, True
```

True = overwrite allowed

---

# 13. File Delete Operations

```vb
If fso.FileExists(filePath) Then
    fso.DeleteFile filePath, True
End If
```

---

# 14. Directory Traversal Rule

Avoid recursive uncontrolled loops.

Bad:

- scanning entire drive
- uncontrolled recursion

Good:

- scoped folder traversal only

---

# 15. Excel File Handling Pattern

## Open → Process → Close

```vb
Dim wb As Workbook

Set wb = Workbooks.Open(filePath)

' process

wb.Close SaveChanges:=False
```

---

# 16. Memory Safety Rule

Always release file objects:

```vb
Set fso = Nothing
Set wb = Nothing
```

---

# 17. Logging File Operations

Every file operation must be logged:

```vb
Utils_Log.Info "FileSystem", "Opening file: " & filePath
```

Errors:

```vb
Utils_Log.Error Err, "FileSystem"
```

---

# 18. File Lock Handling

Detect locked files:

```vb
On Error Resume Next
Set wb = Workbooks.Open(filePath)

If Err.Number <> 0 Then
    Utils_Log.Error Err, "File locked: " & filePath
    Exit Sub
End If
On Error GoTo 0
```

---

# 19. Performance Rules

- avoid repeated file opens
- batch file operations
- minimize disk I/O
- use memory processing (arrays) when possible

---

# 20. Security Rules

- never execute external files automatically
- validate file extensions
- restrict allowed file types:

Allowed:
- `.xlsx`
- `.csv`
- `.txt`

Forbidden:
- `.exe`
- `.bat`
- `.vbs`

---

# 21. AI Rules

When generating file system code, AI must:

- always validate file existence
- never hardcode paths
- use FileSystemObject when possible
- ensure file handles are closed
- sanitize filenames
- log all file operations
- avoid unsafe recursion
- restrict file extensions
- use Environ("TEMP") for temporary files
- ensure robust error handling

---

# Example

## Bad

```vb
Open "C:\data\file.txt" For Output As #1
```

## Good

```vb
Dim ts As Object
Set ts = CreateObject("Scripting.FileSystemObject") _
    .OpenTextFile(filePath, 8, True)

ts.WriteLine "data"
ts.Close
```

---

# Golden Rules

1. Never hardcode file paths.
2. Always validate file existence.
3. Use FileSystemObject.
4. Always close file handles.
5. Log all file operations.
6. Sanitize filenames.
7. Use TEMP for temporary files.
8. Restrict file extensions.
9. Avoid uncontrolled directory traversal.
10. Treat file system as external and unsafe by default.