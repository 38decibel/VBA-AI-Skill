# 00 - AI Principles

**Project:** VBA AI Skill
**Version:** 1.0.0
**Language:** English

---

# Purpose

This document defines the fundamental principles that every AI assistant must follow when generating Microsoft Excel VBA code.

These principles are technology-independent and apply before any coding standard, architecture guideline, or project-specific rule.

The goal is to produce code that is:

* Readable
* Maintainable
* Predictable
* Reusable
* Performant
* Robust
* Easy to review
* Easy to debug
* Safe for production

Whenever a conflict exists between writing shorter code and writing clearer code, clarity always wins.

---

# AI Mission

When generating VBA code, your objective is **not** to produce the shortest possible solution.

Your objective is to produce code that an experienced developer would confidently merge into a production application without major refactoring.

Always assume:

* The project will grow.
* Another developer will maintain the code.
* The code may still exist in ten years.
* The generated code will become part of a larger application.

---

# Primary Objectives

The AI must optimize for the following priorities, in this exact order:

1. Correctness
2. Reliability
3. Readability
4. Maintainability
5. Reusability
6. Performance
7. Conciseness

Never sacrifice a higher priority for a lower one.

---

# Fundamental Philosophy

## Rule 1

Write code for humans.

Source code is read far more often than it is written.

Every line should be immediately understandable.

---

## Rule 2

Explicit is better than implicit.

Never rely on VBA default behavior when an explicit statement improves readability.

Examples:

* Explicit variable types
* Explicit object references
* Explicit parameter passing
* Explicit return values

---

## Rule 3

Readability is more important than cleverness.

Avoid "smart" solutions that reduce clarity.

Simple code is preferred over impressive code.

---

## Rule 4

Maintainability is more important than brevity.

Never reduce the number of lines at the expense of readability.

Longer but clearer code is preferred.

---

## Rule 5

Every procedure should tell a story.

The reader should understand:

* what the procedure does
* why it exists
* how the major steps are organized

without requiring external documentation.

---

# Clean Code Principles

The AI shall follow the following software engineering principles.

## KISS

Keep It Simple, Stupid.

Always choose the simplest solution that correctly solves the problem.

Do not introduce unnecessary abstraction.

---

## DRY

Don't Repeat Yourself.

Duplicate logic must be extracted into reusable procedures.

Avoid copy/paste programming.

---

## YAGNI

You Aren't Gonna Need It.

Do not implement hypothetical future functionality.

Generate only what is required today.

---

## Single Responsibility Principle

Each procedure should perform one logical task.

Each module should represent one domain.

Each class should encapsulate one concept.

---

## Separation of Concerns

Separate:

* Excel interaction
* Business logic
* File system
* User interface
* Configuration
* Logging
* Error handling

Never mix unrelated concerns.

---

# Professional Mindset

Always write production-quality code.

Never generate:

* demo shortcuts
* temporary hacks
* quick fixes
* throwaway code

Every generated solution should be suitable for long-term maintenance.

---

# Predictability

Generated code should behave predictably.

Avoid surprising side effects.

Avoid hidden dependencies.

Avoid modifying global state unless explicitly required.

---

# Consistency

Always use consistent:

* naming
* indentation
* formatting
* comments
* error handling
* object lifecycle

Consistency is more valuable than personal preference.

---

# Defensive Programming

Assume that:

* parameters may be invalid
* worksheets may not exist
* ranges may be empty
* files may be missing
* COM objects may fail
* SAP sessions may disconnect
* users may perform unexpected actions

Validate assumptions before using them.

Never trust external input.

---

# Error Prevention

Prevent errors before handling them.

Prefer validation over exception recovery.

Examples:

* Check whether a worksheet exists.
* Check whether an object is Nothing.
* Check whether a file exists.
* Check whether an array has been initialized.
* Check whether a collection contains an item.

---

# Fail Fast

Detect invalid states as early as possible.

If a required condition is not met:

* stop processing
* log the problem
* return a meaningful result

Do not continue with invalid assumptions.

---

# Code Generation Philosophy

The AI must generate complete code.

Avoid placeholders such as:

```text
' TODO
' Implement here
```

unless explicitly requested.

Generated code should compile whenever sufficient information is available.

---

# Refactoring Philosophy

Prefer multiple small procedures over one large procedure.

If a procedure exceeds approximately fifty lines, recommend extracting helper procedures.

Readable architecture is preferred over monolithic implementations.

---

# Documentation Philosophy

Documentation explains:

* why something exists
* assumptions
* business rules
* design decisions

Documentation should never simply repeat the code.

Bad:

```vb
i = i + 1

' Increment i
```

Good:

```vb
' Skip the header row because it contains column names.
For i = 2 To lastRow
```

---

# Performance Philosophy

Do not optimize prematurely.

Follow this order:

1. Correctness
2. Readability
3. Maintainability
4. Measure performance
5. Optimize bottlenecks

Never sacrifice readability for hypothetical performance improvements.

---

# Simplicity

Prefer:

* clear variables
* intermediate variables
* descriptive names

instead of deeply nested expressions.

Bad:

```vb
If Len(Trim(UCase$(value))) > 0 Then
```

Better:

```vb
cleanValue = Trim$(value)

If Len(cleanValue) > 0 Then
```

---

# Object-Oriented Thinking

Even though VBA is limited, think in terms of objects.

Group related behavior together.

Avoid procedural "God modules."

Prefer:

* classes
* encapsulation
* reusable components

when appropriate.

---

# Design for Reuse

Write code that can be reused.

Avoid hardcoded:

* worksheet names
* workbook names
* file paths
* constants
* business rules

Whenever possible, parameterize behavior.

---

# Review Before Completion

Before considering generated code complete, mentally verify:

* Does it compile?
* Are variables declared?
* Are object references qualified?
* Are objects released?
* Are errors handled?
* Is the naming consistent?
* Is duplication avoided?
* Is the logic obvious?
* Is the code production-ready?

If the answer is "No" to any question, improve the code before presenting it.

---

# AI Behavior Rules

The AI shall always:

* Produce production-quality VBA.
* Prefer maintainability over cleverness.
* Prefer explicitness over implicit behavior.
* Follow established coding standards.
* Use descriptive names.
* Structure code logically.
* Suggest improvements when appropriate.
* Explain important design decisions.
* Recommend refactoring when beneficial.

---

# AI Shall Never

The AI shall never:

* Generate code without `Option Explicit`.
* Depend on `ActiveWorkbook` unless explicitly requested.
* Depend on `ActiveSheet` unless explicitly requested.
* Depend on `Selection` unless absolutely necessary.
* Depend on `Select` or `Activate` for automation.
* Leave duplicated code.
* Ignore errors silently.
* Generate obsolete VBA practices.
* Introduce hidden side effects.
* Use magic numbers without explanation.
* Prefer cleverness over readability.
* Recommend global variables when local scope is sufficient.
* Mix UI logic with business logic.
* Generate partially implemented code without warning.

---

# Definition of Success

Generated VBA code is considered successful only if it satisfies all of the following:

* Correct
* Readable
* Maintainable
* Reusable
* Well documented
* Fully typed
* Properly structured
* Robust against failures
* Easy to review
* Easy to extend
* Consistent with professional software engineering standards

If any of these characteristics is missing, the generated solution should be improved before delivery.

---

**End of Document**
