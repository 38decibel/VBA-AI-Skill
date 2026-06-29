# 01 - General Philosophy (VBA AI Skill Framework)

## Objective

This document defines the **overall philosophy for building VBA applications** in this framework.

It translates the core AI principles into **practical development mindset rules** for real projects.

---

# Golden Rule

> VBA code must be designed as a small, predictable, maintainable system — not a collection of macros.

---

# 1. Think in Systems, Not Macros

Traditional VBA development often produces isolated macros.

This framework enforces:

- modular architecture
- reusable components
- clear separation of responsibilities

---

## ❌ Old mindset

- “one macro per task”
- copy/paste logic
- spreadsheet-centric thinking

---

## ✔ Target mindset

- “one system composed of layers”
- reusable services
- structured workflows

---

# 2. Excel is the Interface, Not the Core

Excel is only:

- input surface
- output renderer
- user interaction layer

It is NOT:

- a computation engine
- a business logic container
- a state manager

---

## Rule

> All logic must be executed outside of worksheet cells.

---

# 3. Simplicity is a Design Requirement

Complexity is treated as a defect.

AI must prioritize:

- simple flows
- clear structures
- minimal dependencies

---

## Anti-goal

- not “most clever solution”
- not “shortest code”
- not “advanced VBA tricks”

---

## Goal

- readable in 30 seconds
- predictable behavior
- easy to modify safely

---

# 4. Predictability Over Flexibility

A system that behaves unpredictably is considered broken.

AI must ensure:

- deterministic execution
- explicit inputs/outputs
- no hidden side effects

---

# 5. Structure Over Speed of Development

Short-term speed is NOT a priority.

Instead:

- enforce architecture
- enforce patterns
- enforce consistency

This reduces long-term cost.

---

# 6. Everything is a Component

Every part of the system must belong to a layer:

- UserForms → presentation layer
- Modules → orchestration layer
- Classes → business logic layer
- Utils → shared infrastructure

---

# 7. Code Must Tell the Story

Good VBA code should read like:

> “What is happening?” not “How is it implemented?”

---

## Example

### Bad

```vb
If ws.Cells(i, 1).Value <> "" Then
    ws.Cells(i, 2).Value = ws.Cells(i, 1).Value * 2
End If
```

---

### Good

```vb
If order.IsValid Then
    order.CalculateTotal
End If
```

---

# 8. Excel Should Be Treated as Volatile

AI must assume:

- users modify sheets
- structure changes
- data is incomplete
- formats are inconsistent

Therefore:

- always validate
- never trust worksheet structure blindly

---

# 9. Code Must Be Easy to Debug

Debugging is a first-class requirement.

Therefore:

- structured logging is mandatory
- procedures must be small
- execution flow must be traceable

---

# 10. Avoid “Spaghetti VBA”

Spaghetti VBA is defined as:

- logic spread across modules
- hidden dependencies
- uncontrolled global state
- circular calls

AI must actively prevent this.

---

# 11. Reusability is Mandatory

If logic appears twice → it must be extracted.

If logic is unclear → it must be simplified.

---

# 12. Naming Reflects Intent

Names must describe:

- business meaning
- not technical implementation

---

## Examples

### Bad

- `ProcessData`
- `DoStuff`
- `HandleSheet`

### Good

- `GenerateInvoice`
- `ValidateOrder`
- `ExportZplLabel`

---

# 13. VBA is Not an Island

VBA systems must integrate with:

- Excel
- files
- COM applications
- external systems

Therefore:

> Every external dependency is unreliable.

---

# 14. Failure is Normal

AI must assume:

- files fail
- COM fails
- data is missing
- users make mistakes

So:

- failures must be expected
- not exceptional

---

# 15. Maintainability is the Primary KPI

Not:

- speed of writing code
- number of lines saved
- clever tricks

But:

- ease of modification
- readability
- stability over time

---

# 16. Design for the Next Developer

Every piece of code should assume:

> someone else will maintain it in 6 months.

That “someone” must understand it quickly.

---

# 17. Minimal Cognitive Load

AI must ensure:

- low nesting
- short procedures
- clear naming
- predictable structure

---

# 18. Avoid Over-Abstraction

Abstraction is useful only when:

- duplication exists
- complexity increases
- reuse is real

Otherwise:

> simple code is better than abstract code.

---

# 19. Excel Automation is Industrial Logic

This framework is designed for:

- business automation
- data processing pipelines
- enterprise Excel systems

Not:

- quick personal macros
- experimental scripts

---

# 20. Philosophy Summary

This framework is built on:

1. Simplicity over cleverness
2. Structure over speed
3. Predictability over flexibility
4. Maintainability over brevity
5. Safety over convenience
6. Clarity over optimization
7. Explicitness over implicit behavior

---

# Golden Statement

> VBA code is successful only when it becomes invisible in operation but obvious in structure.