# Code Review Checklist — VBA AI Skill + SAP Tools & Automations

## Objective

This checklist ensures every VBA change is correct, performant, maintainable, aligned with
architecture rules, and safe for production. It merges the generic VBA AI Skill standard with
the project-specific rules for SAP GUI Scripting automation (SAP Tools & Automations).

It is used for both human review and AI self-validation before code is considered complete.

---

## Golden Rule

> If code does not pass this checklist, it is not production-ready.

---

## 1. Module & Header Compliance

- [ ] Module header present: `@Module / @Description / @Dependencies / @Author / @Version / @Updated`
- [ ] Version follows `MAJOR.MINOR.PATCH`
- [ ] `Option Explicit` present
- [ ] `Option Private Module` present (standard modules)
- [ ] One module = one responsibility (no "God modules")

---

## 2. Architecture & Separation of Concerns

- [ ] UserForms contain UI logic only (no business logic, no SAP logic)
- [ ] Modules (`SAP_*`) contain orchestration / high-level action sequences only
- [ ] `SAP_Interact` contains low-level SAP interaction only (FindById, FillBox, SendKey)
- [ ] Class modules (`cls*`) encapsulate one concept; no business logic in `clsSapEvents` beyond dispatching
- [ ] `Utils_*` modules contain shared, generic utilities only — no SAP- or business-specific logic
- [ ] No UI logic mixed with SAP logic in the same function
- [ ] No logic leakage: no Excel manipulation inside event handlers, no business logic in events

---

## 3. Naming Conventions

- [ ] Module-level variables prefixed `m_`
- [ ] Global/Public variables prefixed `g_`
- [ ] Local constants prefixed `C_` (Private); shared constants live in `Constants` module
- [ ] Functions/Subs in PascalCase, verb + noun, intent-revealing (no `DoStuff`, `ProcessData`)
- [ ] Boolean functions prefixed `Is` / `Has`
- [ ] Getters prefixed `Get`
- [ ] Event handlers prefixed `On`, kept as thin wrappers (delegate to a real function)
- [ ] Classes prefixed `cls`
- [ ] Modules in PascalCase with domain prefix (`SAP_`, `Utils_`)

---

## 4. Comments & Documentation

- [ ] Public function header present (description, parameters, return value)
- [ ] Comments in English only — no French, no mixed language
- [ ] No accented or special characters in comments
- [ ] Comments explain WHY, not WHAT
- [ ] At least one comment per logical block (~5–10 lines), blank line before the block

---

## 5. Function Design

- [ ] One function = one responsibility
- [ ] Max ~50 lines per function; longer functions flagged and split
- [ ] SAP-interacting functions return `Boolean` (success/failure) or `Object` (result/`Nothing`) — never `String`
- [ ] No magic numbers — use `Constants` or named variables
- [ ] Guard clauses used for early exit instead of deep nesting

---

## 6. Error Handling

- [ ] `On Error GoTo ErrHandler` present at the top of every Public function
- [ ] `Cleanup:` label present, used to release all object references
- [ ] `ErrHandler:` label present, calls `LogError` with format `"ModuleName.FunctionName: description"`
- [ ] `On Error Resume Next` limited to 2–3 lines max, then restored with `On Error GoTo ErrHandler`
- [ ] No error swallowed silently — every catch path logs
- [ ] Errors re-raised to caller when appropriate

---

## 7. Logging

- [ ] `LogDebug` / `LogInfo` / `LogError` used — never bare `Debug.Print`, never `MsgBox` for debugging
- [ ] Format respected: `[LEVEL] HH:MM:SS | Module.Function | Message`
- [ ] `LogDebug` at function entry for key parameters
- [ ] `LogInfo` for meaningful results
- [ ] `LogError` present in every `ErrHandler`
- [ ] Logging does not degrade performance inside loops

---

## 8. Object Lifecycle (COM & General)

- [ ] `Set obj = Nothing` in `Cleanup` for every object declared in the function
- [ ] No SAP COM object reference stored beyond function scope (except `m_session` in `SAP_Logon`)
- [ ] Object use guarded with `If Not obj Is Nothing Then`
- [ ] Late binding used for COM unless early binding is explicitly justified
- [ ] No COM calls inside performance-critical loops
- [ ] After `CloseSession`, the session variable is set to `Nothing`

---

## 9. SAP GUI Scripting / COM Automation Rules

- [ ] Use the SAP GUI Script Recorder before coding any new navigation sequence (never guess `btn[n]` / `SendVKey`)
- [ ] New-session pattern followed exactly:
      `SaveCurrentSessionRef` → `PrepareNewSessionCapture` → `SendCommand("/oTRX")` →
      `Iconify` → `CaptureNewSession` → `UseSession` → [work] →
      `ClearSessionOverride` → `RestoreSessionRef` → `CloseSession`
- [ ] `Iconify` precedes `CaptureNewSession`, not the reverse
- [ ] Cleanup sequencing respected: session restore happens before object release
- [ ] Popup windows (`wnd[1]`) detected conditionally, not assumed present
- [ ] Functions that depend on window/title state (e.g. `ExtractTotalItemsFromTitle`) are called only after navigation completes
- [ ] `newSes.FindById("wnd[0]").Close` called in `Cleanup` to close the opened session
- [ ] No SAP session used without `If Not ses Is Nothing` guard
- [ ] Statusbar message type checked before acting on result (`g_SB_SUCCESS`, `g_SB_ERROR`, etc.)

---

## 10. Excel Object Model Usage

- [ ] No `ActiveWorkbook` / `ActiveSheet` / `Selection` / `Select` / `Activate` in business or SAP logic
- [ ] Explicit worksheet/workbook references used (`wsDashboard`, not `Sheets("Dashboard")`)
- [ ] `ListObject` used for structured data; constants store table names (e.g. `g_TABLE_MB90_DOCS`)
- [ ] `.Value2` used for bulk data read/write operations
- [ ] `LastRow` / bounds computed dynamically, never hardcoded

---

## 11. Performance

- [ ] No premature optimization — correctness and readability come first
- [ ] No cell-by-cell loops over large datasets; arrays used for bulk operations
- [ ] Excel I/O minimized inside loops
- [ ] `Application.ScreenUpdating` / `Calculation` / `EnableEvents` disabled around heavy operations and restored after
- [ ] SAP round-trips minimized/batched where possible
- [ ] `Dictionary` used for lookups, not linear search

---

## 12. Data Structures

- [ ] `Scripting.Dictionary` used for structured lookups and data passing between modules
- [ ] Collections reserved for simple ordered lists
- [ ] Arrays validated (`IsEmpty` / bounds check) before use
- [ ] `ReDim Preserve` used sparingly, never inside hot loops
- [ ] Dictionary key existence checked with `.Exists` before access

---

## 13. Defensive Programming

- [ ] All inputs validated at function entry
- [ ] Objects checked for `Nothing` before use
- [ ] File existence verified before file access; file handles closed properly
- [ ] No hardcoded file paths — all paths come from `Constants`
- [ ] No implicit assumptions about SAP session/screen state

---

## 14. Windows API

- [ ] API declarations isolated in `Utils_Windows` only
- [ ] API-related constants stored in `Constants` (e.g. `g_SW_MAXIMIZE`)
- [ ] Return values of critical API calls validated
- [ ] No raw API calls scattered in business/SAP logic

---

## 15. Stability & Safety

- [ ] No uncontrolled recursion or infinite loops
- [ ] No unbounded range/array operations
- [ ] No unsafe type assumptions (explicit typing throughout, `Variant` only when justified)
- [ ] Reentrancy considered for event handlers (`EnableEvents` managed safely)

---

## 16. Refactoring Quality

- [ ] No duplicated logic (DRY) — shared logic extracted into reusable helpers
- [ ] Functions over ~50 lines split into sub-functions
- [ ] Naming revisited if it no longer reflects actual responsibility
- [ ] No accidental behavior change introduced by refactor

---

## 17. AI-Generated Code Specific Checks

- [ ] No `' TODO` / placeholder left without explicit justification
- [ ] Code compiles given the information available
- [ ] Important design decisions explained in comments
- [ ] Refactor suggested proactively if a function exceeds ~50 lines
- [ ] Rule hierarchy respected: project-specific rules > VBA AI Skill > MS recommendations > general best practices

---

## 18. Red Flags (Automatic Rejection)

Reject the code if it contains any of:

- `ActiveSheet` / `ActiveWorkbook` / `Select` / `Activate` in logic layers
- Cell-by-cell loops on large datasets
- Missing error handling on a Public function
- Business or SAP logic inside a UserForm
- COM objects not released in `Cleanup`
- Silent error suppression (`On Error Resume Next` with no recovery/log)
- New SAP session opened without the full session-management pattern
- Hardcoded SAP `FindById` paths repeated without a constant
- French or accented comments

---

## 19. Final Review Question

> "Would I trust this code in production, driving a real SAP session with real data, unsupervised?"

If the answer is no → refactor required before merge.

---

## Golden Rules Summary

1. Architecture and separation of concerns must be respected.
2. Every Public function has structured error handling and a `Cleanup` block.
3. Every SAP session follows the canonical save/capture/restore pattern.
4. Excel and SAP are both treated as unreliable I/O — validate, don't assume.
5. No magic numbers, no hardcoded paths — everything lives in `Constants`.
6. Logging must be consistent (`LogDebug`/`LogInfo`/`LogError`), never `Debug.Print` or `MsgBox`.
7. No business logic in UI, events, or `clsSapEvents`.
8. Functions stay under ~50 lines and do one thing.
9. Comments are in English, unaccented, and explain WHY.
10. Production readiness is the only acceptable standard.
