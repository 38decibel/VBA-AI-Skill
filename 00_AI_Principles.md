# 00 - AI Principles (VBA AI Skill Foundation)

## Objective

This document defines the **core principles that govern all VBA code generation and design decisions** in this framework.

It is the highest-level rule set: every other guideline is derived from it.

---

# Golden Rule

> AI must prioritize correctness, safety, and maintainability over cleverness or brevity.

---

# 1. Determinism First

VBA automation must be predictable.

AI must ensure:

- same input → same output
- no hidden state dependencies
- no reliance on UI state (ActiveSheet, Selection, etc.)

---

# 2. Excel is a Data Layer, Not a Processor

Excel must be treated as:

> a storage and presentation system, not a computation engine.

Therefore:

- logic must be executed in VBA memory
- Excel is only used for input/output
- avoid cell-by-cell computation

---

# 3. Safety Over Convenience

AI must assume:

- files may be missing
- data may be corrupted
- COM objects may fail
- users may input invalid values

Therefore:

- validate everything
- fail gracefully
- never trust external input

---

# 4. Explicitness Over Implicitness

AI must always prefer:

- explicit variables over implicit ones
- explicit references over global state
- explicit control flow over hidden behavior

---

## Forbidden implicit behaviors

- ActiveSheet usage
- implicit type conversion without checks
- hidden dependencies between modules

---

# 5. Performance Awareness

AI must design code with:

- bulk operations (arrays)
- minimal Excel interaction
- minimized COM calls
- reduced loops over worksheet cells

Performance is a **design constraint**, not an optimization step.

---

# 6. Structured Architecture First

All generated code must respect layered architecture:

```
UI Layer (UserForms)
    ↓
Controller Layer
    ↓
Business Logic (Classes)
    ↓
Data Layer (Excel / Files / COM)
```

No exceptions.

---

# 7. Separation of Concerns

Each component must have a single responsibility:

- UserForms → UI only
- Modules → orchestration only
- Classes → business logic
- Utils → shared services (logging, helpers)

---

# 8. Defensive Mindset

AI must assume:

> everything outside VBA memory is unreliable.

Therefore:

- validate inputs early
- check object existence
- handle empty ranges
- protect against type mismatch
- log all failures

---

# 9. Fail Loud, Fail Clearly

Errors must:

- never be silent
- always be logged
- include context (module, procedure, key variables)

---

# 10. Logging is Mandatory

Every meaningful operation must include:

- Debug logs for development
- Info logs for business flow
- Error logs for failures

Logging is not optional.

---

# 11. No UI Dependency in Logic

Business logic must NEVER depend on:

- MsgBox
- InputBox
- UserForms state
- ActiveWorkbook / ActiveSheet

---

# 12. Minimize Excel Object Interaction

AI must:

- avoid repeated `.Cells` access
- avoid loops over ranges
- batch read/write with `.Value2`
- cache data in memory

---

# 13. Consistency Over Flexibility

AI must prioritize:

- standardized patterns
- reusable templates
- predictable structure

Even if alternative solutions exist.

---

# 14. Code Must Be Readable Without Comments

If code requires excessive comments to be understood:

→ it must be refactored.

---

# 15. No Hidden State

AI must avoid:

- global variables
- implicit module-level dependencies
- hidden state transitions

State must always be explicit and controlled.

---

# 16. COM, API, and External Systems Are Untrusted

All external systems must be treated as:

- unstable
- slow
- failure-prone

Therefore:

- always wrap COM calls
- always handle errors
- always release resources

---

# 17. Simplicity Over Cleverness

AI must prefer:

- simple loops over complex abstractions
- clear logic over optimization tricks
- explicit code over advanced VBA hacks

---

# 18. Refactoring is Part of Design

AI must assume:

> code will evolve.

Therefore:

- modular structure is mandatory
- duplication must be avoided
- long procedures must be split

---

# 19. No Premature Optimization

AI must:

- first ensure correctness
- then ensure readability
- then optimize performance

Never optimize before clarity.

---

# 20. Production-Ready by Default

AI output must assume:

- real Excel data
- real users
- real failures
- real integrations

No “demo code”.

---

# 21. AI Responsibility Rule

AI is responsible for:

- architecture compliance
- performance safety
- error handling completeness
- logging consistency
- maintainability of generated code

---

# 22. Output Quality Standard

Generated code must:

- compile without modification
- follow naming conventions
- respect module boundaries
- include error handling
- include logging
- avoid anti-patterns

---

# 23. Evolution Principle

This framework is designed to evolve:

- rules may be refined
- patterns may be optimized
- templates may be extended

But core principles remain stable.

---

# Golden Summary

1. Excel is I/O, not a processor.
2. Safety is mandatory.
3. Explicit beats implicit.
4. Architecture is non-negotiable.
5. Performance is a design constraint.
6. Logging is always required.
7. External systems are untrusted.
8. Simplicity wins over cleverness.
9. Refactoring is continuous.
10. Production readiness is the default.