# 24 - Code Review Checklist (VBA AI Skill Framework)

## Objective

This checklist ensures every VBA change is:

- correct
- performant
- maintainable
- aligned with architecture rules
- safe for production

It is used for **human review and AI validation**.

---

# Golden Rule

> If code does not pass this checklist, it is not production-ready.

---

# 1. Architecture Compliance

## ✔ Separation of concerns

- [ ] UserForms contain only UI logic
- [ ] Modules contain orchestration only
- [ ] Class Modules contain business logic
- [ ] Utils modules contain shared utilities only

---

## ✔ No logic leakage

- [ ] No business logic in UserForms
- [ ] No Excel manipulation in events
- [ ] No UI logic in classes

---

# 2. Naming Quality

- [ ] Variables are meaningful
- [ ] Procedures describe intent (verb + noun)
- [ ] Classes represent business entities
- [ ] No vague names (`DoStuff`, `ProcessData`)

---

# 3. Excel Object Model Usage

- [ ] No `ActiveSheet` usage in business logic
- [ ] No `Select` or `Activate`
- [ ] Explicit worksheet references used
- [ ] ListObjects used for structured data
- [ ] `.Value2` used for data operations

---

# 4. Performance Checks

- [ ] No cell-by-cell loops over large datasets
- [ ] Arrays used for bulk operations
- [ ] Excel access minimized inside loops
- [ ] ScreenUpdating disabled for heavy operations
- [ ] Events disabled during batch updates
- [ ] Calculation mode handled properly

---

# 5. Data Structures

- [ ] Arrays used for Excel data processing
- [ ] Dictionary used for lookups
- [ ] Collections only for simple lists
- [ ] No misuse of ReDim in loops

---

# 6. Error Handling

- [ ] Every procedure has error handler
- [ ] No silent `On Error Resume Next` misuse
- [ ] Errors logged via `Utils_Log.Error`
- [ ] Error context includes module/procedure name

---

# 7. Logging Standards

- [ ] Key actions logged (Info/Debug)
- [ ] Errors logged with context
- [ ] No MsgBox used for debugging
- [ ] Logging does not break performance loops

---

# 8. Defensive Programming

- [ ] Inputs validated early
- [ ] Objects checked for Nothing
- [ ] Arrays validated before use
- [ ] Dictionary keys checked with `.Exists`
- [ ] File existence verified before access

---

# 9. UserForm Rules

- [ ] No business logic in forms
- [ ] Events are minimal
- [ ] Calls delegated to controllers
- [ ] UI state is separated from data logic

---

# 10. Event Handling Rules

- [ ] Events are thin wrappers only
- [ ] No loops in event handlers
- [ ] EnableEvents managed safely
- [ ] Reentrancy considered

---

# 11. COM Automation Rules

- [ ] COM objects released properly
- [ ] Late binding used unless justified
- [ ] No COM in performance-critical loops
- [ ] Error handling present for external apps

---

# 12. File System Rules

- [ ] No hardcoded paths
- [ ] File existence checked
- [ ] File handles closed properly
- [ ] TEMP folder used when appropriate
- [ ] Safe file naming applied

---

# 13. Windows API Rules

- [ ] API calls wrapped in services
- [ ] 32/64-bit compatibility ensured
- [ ] No raw API calls in business logic
- [ ] Return values validated

---

# 14. Refactoring Quality

- [ ] No duplicated logic
- [ ] Long procedures split
- [ ] Guard clauses used instead of nesting
- [ ] Naming improved during refactor
- [ ] No accidental behavior changes

---

# 15. Maintainability

- [ ] Code is readable without comments
- [ ] Logic is modular
- [ ] Dependencies are explicit
- [ ] No hidden global state

---

# 16. Stability & Safety

- [ ] No uncontrolled recursion
- [ ] No infinite loops
- [ ] No unbounded range operations
- [ ] No unsafe type assumptions

---

# 17. Performance Validation

- [ ] Bulk operations used
- [ ] Excel I/O minimized
- [ ] Dictionary used for fast lookups
- [ ] Arrays used instead of ranges in loops

---

# 18. AI Code Generation Rules

When reviewing AI-generated VBA code, ensure:

- [ ] strict adherence to architecture layers
- [ ] no procedural dumping logic
- [ ] proper abstraction usage
- [ ] performance-first design
- [ ] defensive programming applied

---

# 19. Red Flags (Automatic Rejection)

Reject code if it contains:

- ActiveSheet usage in logic layer
- Select / Activate
- Cell-by-cell loops on large data
- Missing error handling
- Business logic in UserForms
- COM objects not released
- Silent error suppression

---

# 20. Final Review Question

Before approving code, ask:

> “Would I trust this code in production with real Excel data and no supervision?”

If the answer is no → refactor required.

---

# Golden Rules

1. Architecture must be respected.
2. Performance must be intentional.
3. Error handling is mandatory.
4. Excel must be treated as I/O only.
5. Code must be readable and modular.
6. No business logic in UI or events.
7. All external systems are unreliable.
8. Logging must be consistent.
9. No hidden state allowed.
10. Production readiness is the only standard.