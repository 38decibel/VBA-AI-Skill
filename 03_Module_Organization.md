03 - Module Organization

Project: VBA AI Skill
Chapter: 03 - Module Organization
Version: 1.0.0

---

Purpose

This chapter defines the standards for organizing modules within a Microsoft Excel VBA project.

A well-organized project improves readability, maintainability, scalability, and code discoverability.

The objective is to make every module immediately understandable, even to a developer unfamiliar with the project.

---

Philosophy

Modules are the primary building blocks of a VBA application.

Each module should represent a single domain, service, or responsibility.

A developer should be able to understand the purpose of a module simply by reading its name.

---

One Module = One Responsibility

Each module shall implement one logical concern.

Examples:

Good

Excel_Table
Excel_Worksheet
Excel_Formatting

Utils_Array
Utils_File
Utils_String

Business_Orders
Business_Customers
Business_Validation

Logging

Configuration

SAP_MB03
SAP_MB90

Bad

Helpers
Misc
Functions
General
Module1
Utilities
Various

Generic modules inevitably become dumping grounds.

---

Maximum Module Size

A module should remain reasonably small.

Recommended limits:

Metric| Recommendation
Public procedures| < 20
Total procedures| < 40
Code size| Preferably under 1000 lines

When a module becomes difficult to navigate, split it.

---

Domain-Based Organization

Modules should be grouped by domain rather than by technical implementation.

Good

Business_Orders
Business_Invoices
Business_Products

Instead of

Functions1
Functions2
Functions3

---

Module Naming

Module names should use PascalCase.

Words should be separated by underscores.

Examples

Constants
Configuration

Logging

Utils_Array
Utils_String
Utils_File

Excel_Table
Excel_Workbook

SAP_MB03
SAP_MM03
SAP_Logon

Business_Validation
Business_Report

---

Module Prefixes

The following prefixes are recommended.

Business

Contains business rules.

Business_

Examples

Business_Orders
Business_Validation
Business_Pricing

---

Excel

Contains Excel-specific operations.

Excel_

Examples

Excel_Table
Excel_Worksheet
Excel_Formatting

---

Utils

Contains reusable generic helpers.

Utils_

Examples

Utils_Array
Utils_File
Utils_Date
Utils_String

---

Config

Application configuration.

Configuration

or

Config

---

Logging

Logging service.

Logging

or

Utils_Logging

depending on project size.

---

External Systems

External integrations should be grouped.

Examples

SAP_
Outlook_
Word_
REST_
SQL_

---

Module Header

Every standard module shall begin with a documentation header.

Example

' =============================================================================
'@Module      Excel_Table
'@Description Excel table helper functions.
'
'@Dependencies
'   Constants
'   Logging
'
'@Author      Daniel Moindrot
'@Version     1.0.0
'@Updated     YYYY-MM-DD
' =============================================================================

Option Explicit
Option Private Module

The header documents ownership and responsibility.

---

Option Explicit

Every module must include

Option Explicit

without exception.

---

Option Private Module

Unless a module intentionally exposes procedures to other VBA projects, use

Option Private Module

This reduces the public API surface.

---

Procedure Ordering

Procedures should follow a consistent order.

Recommended order:

1. Module constants
2. Module variables
3. Enumerations
4. Public Types
5. Public Functions
6. Public Subs
7. Private Functions
8. Private Subs

This makes navigation predictable.

---

Constants

Module-specific constants belong at the beginning of the module.

Example

Private Const C_DEFAULT_TIMEOUT As Long = 5000

Application-wide constants belong in the Constants module.

---

Module Variables

Declare module variables immediately after constants.

Example

Private m_logger As clsLogger

Private m_isInitialized As Boolean

Avoid scattered declarations.

---

Public Interface

The public procedures define the module interface.

Keep the public interface small.

Expose only procedures intended for external use.

Everything else should be Private.

---

Private Helpers

Complex public procedures should delegate work to private helper procedures.

Example

ReadDocument()

↓

ValidateInput()

↓

ReadHeader()

↓

ReadItems()

↓

BuildResult()

↓

Cleanup()

Private procedures improve readability.

---

Logical Sections

Large modules should be divided into logical sections.

Example

Constants

Module Variables

Initialization

Public Interface

Validation

Business Logic

Private Helpers

Cleanup

---

Avoid Circular Dependencies

Modules should not depend on each other recursively.

Avoid

Module A

↓

Module B

↓

Module C

↓

Module A

Circular dependencies complicate maintenance.

---

Avoid God Modules

A God Module performs unrelated tasks.

Bad

Utils

Contains

Files

Dates

Strings

Logging

SAP

Excel

JSON

Printing

Good

Utils_File

Utils_Date

Utils_String

Utils_JSON

Small focused modules are easier to maintain.

---

Cross-Domain Calls

Business modules should not directly manipulate Excel whenever possible.

Instead

Business

↓

Excel

rather than

Excel

↓

Business

↓

Excel

---

Internal Dependencies

Prefer this dependency flow.

Business

↓

Services

↓

Utilities

↓

Infrastructure

Never reverse this dependency without justification.

---

Reusability

Modules should be reusable.

Avoid embedding project-specific assumptions inside generic utility modules.

Example

Bad

Public Function GetSAPMaterial()

inside

Utils_String

Good

SAP_MM03

---

Initialization

If initialization is required, expose a dedicated Initialize procedure.

Avoid implicit initialization whenever possible.

Example

Public Sub Initialize()

End Sub

---

Cleanup

Modules managing resources should expose cleanup procedures when appropriate.

Example

Public Sub Shutdown()

End Sub

---

Documentation

Every public procedure should include a documentation header.

Every logical block should include meaningful comments explaining why the code exists.

Avoid commenting obvious statements.

---

Module Cohesion

Everything inside a module should contribute to the same responsibility.

If a procedure feels "out of place", it probably belongs elsewhere.

---

AI Module Generation Rules

When generating new modules, the AI shall:

- Create focused modules.
- Avoid generic names.
- Include a module header.
- Use Option Explicit.
- Use Option Private Module unless otherwise required.
- Keep the public interface minimal.
- Organize procedures consistently.
- Separate public and private procedures.
- Extract helper procedures when complexity increases.
- Avoid oversized modules.
- Recommend splitting large modules.

---

AI Review Checklist

Before generating a module, verify:

- Does the module have one responsibility?
- Is the name explicit?
- Are unrelated procedures absent?
- Is the public API minimal?
- Is the header complete?
- Are declarations organized?
- Is the module scalable?
- Would another developer immediately understand its purpose?

If any answer is negative, reorganize the module before presenting it.

---

Guiding Principle

«A module should represent a single domain of responsibility. If its purpose cannot be described in one sentence, it probably does too much.»

Well-organized modules reduce maintenance costs, improve discoverability, and make large VBA applications easier to evolve over time.

---

End of Chapter 03