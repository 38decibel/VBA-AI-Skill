# 19 - Windows API (Advanced VBA Interoperability)

## Objective

Windows API calls allow VBA to extend beyond Excel limitations:

- file dialogs (advanced)
- window manipulation
- performance timing
- system information
- external process interaction

However, API usage is **low-level, unsafe by default, and platform-sensitive**.

This chapter defines strict rules for safe API integration.

---

# Golden Rule

> Windows API is a last-resort tool. Prefer native VBA or Excel objects whenever possible.

---

# 1. When to Use Windows API

Use API only when:

- VBA has no native equivalent
- performance requires system-level access
- advanced UI/system behavior is needed
- automation cannot be achieved otherwise

---

## Valid use cases

- high-resolution timers
- advanced file dialogs
- window focus management
- keyboard/mouse hooks (rare)
- system-level information

---

## Forbidden use cases

- business logic
- Excel data processing
- replacing ListObject / Range logic
- UI forms logic that VBA already supports

---

# 2. Declaration Rule

All API declarations must be:

- in a dedicated module (e.g. `modWinAPI`)
- clearly grouped
- commented with purpose

---

## Example

```vb
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
```

---

# 3. 64-bit Compatibility Rule

All APIs must support:

- 32-bit VBA
- 64-bit VBA (PtrSafe mandatory)

---

## Rule

> Every API declaration must include conditional compilation.

---

# 4. No Inline API Declarations

Bad:

```vb
Private Declare Function GetTickCount Lib "kernel32" () As Long
```

inside a class or random module.

Good:

- centralized module only

---

# 5. Naming Convention

API wrappers must be wrapped in VBA-friendly functions.

## Bad (raw API usage)

```vb
Dim t As Long
t = GetTickCount()
```

## Good (wrapped)

```vb
Public Function SystemTimeMs() As Long
    SystemTimeMs = GetTickCount()
End Function
```

---

# 6. Never Expose Raw API to Business Logic

Business code must NEVER call API directly.

Bad:

```vb
Call GetTickCount()
```

Good:

```vb
Call TimerService.GetElapsedTime()
```

---

# 7. Wrapper Pattern (Mandatory)

All API calls must be wrapped in service modules:

Examples:

- `WinApi_WindowService`
- `WinApi_FileDialogService`
- `WinApi_TimerService`

---

## Example wrapper

```vb
Public Function GetTimeMs() As Long
    GetTimeMs = GetTickCount
End Function
```

---

# 8. Error Handling Rule

Windows API does NOT throw VBA errors reliably.

You must:

- validate return values manually
- check zero / null results
- log failures explicitly

Example:

```vb
Dim result As Long
result = GetTickCount()

If result = 0 Then
    Utils_Log.Error "API", "GetTickCount failed"
End If
```

---

# 9. Memory Safety Rule

API calls may:

- leak handles
- lock resources
- crash Excel if misused

Always:

- release handles
- avoid pointer misuse
- validate inputs

---

# 10. Pointer Safety (Strict Rule)

Pointer usage is **restricted**.

Allowed only if:

- absolutely necessary
- fully encapsulated
- thoroughly tested

Otherwise:

> Do NOT use AddressOf, CopyMemory, or manual pointers.

---

# 11. File Dialog API Pattern

Preferred modern alternative:

```vb
Application.GetOpenFilename
```

Only use Windows API if:

- advanced customization required

---

# 12. Timing API Pattern

Instead of:

```vb
GetTickCount
```

Use wrapper:

```vb
TimerService.Start
TimerService.ElapsedMs
```

---

# 13. Window Management Rules

Allowed:

- bring Excel to front
- minimize/maximize window

Forbidden:

- manipulating other applications without strict need
- background UI hooks

---

# 14. Debugging API Calls

Always log API interactions:

```vb
Utils_Log.Debug "API", "Calling GetTickCount"
```

---

# 15. Performance Considerations

API is fast but:

- overhead exists for repeated calls
- improper usage can slow Excel
- prefer batching logic

---

# 16. Security Rules

Windows API can:

- access system resources
- manipulate OS behavior

Therefore:

- never execute external code
- never load unknown DLLs
- only use trusted system libraries (kernel32, user32)

---

# 17. Architecture Rule

Windows API must be isolated:

```
Business Layer
   ↓
Service Layer (Wrappers)
   ↓
WinAPI Module
   ↓
Windows OS
```

---

# 18. AI Rules

When generating VBA code, AI must:

- never call Windows API directly in business logic
- always wrap API calls in service modules
- ensure 32/64-bit compatibility
- avoid pointer-based APIs unless explicitly required
- validate API return values
- log all API usage via Utils_Log
- avoid exposing raw API functions
- centralize declarations in one module
- prefer native VBA alternatives first

---

# Example

## Bad

```vb
MsgBox GetTickCount()
```

---

## Good

```vb
Dim ms As Long
ms = TimerService.GetTimeMs()

Utils_Log.Debug "Timer", "Elapsed=" & ms
```

---

# Golden Rules

1. Windows API is a last resort.
2. Never use API in business logic.
3. Always wrap API calls in services.
4. Ensure 32/64-bit compatibility.
5. Never expose raw API functions.
6. Validate all return values.
7. Centralize all declarations.
8. Prefer VBA native functions first.
9. Log all API usage.
10. Treat Windows API as unsafe by default.