02 - Project Architecture

Project: VBA AI Skill
Chapter: 02 - Project Architecture
Version: 1.0.0

---

Purpose

This chapter defines the recommended architecture for professional Microsoft Excel VBA applications.

The objective is to produce applications that are modular, scalable, maintainable and easy to extend.

These principles apply regardless of the business domain.

---

Architecture Philosophy

An Excel workbook is not the application.

The workbook is merely the host of the application.

The VBA project contains the application itself.

The workbook should contain as little business logic as possible.

---

Architectural Objectives

A well-designed VBA application should be:

- Modular
- Easy to understand
- Easy to maintain
- Easy to test
- Easy to extend
- Resistant to change
- Reusable across projects

---

Layered Architecture

Whenever possible, organize the application into logical layers.

+------------------------------------------------+
|                 User Interface                 |
|  Worksheets / Ribbon / UserForms / Buttons     |
+------------------------------------------------+
                    |
                    v
+------------------------------------------------+
|            Application Services                |
| Workflow - Orchestration - Controllers         |
+------------------------------------------------+
                    |
                    v
+------------------------------------------------+
|               Business Logic                   |
| Validation - Calculations - Rules             |
+------------------------------------------------+
                    |
                    v
+------------------------------------------------+
|            Infrastructure Layer                |
| Excel - Files - COM - SAP - API - Logging     |
+------------------------------------------------+

Each layer should have a clearly defined responsibility.

---

Separation of Concerns

Never mix unrelated responsibilities.

Separate:

- User Interface
- Business Logic
- Excel Interaction
- File System
- Logging
- Configuration
- External Systems
- Utilities

A procedure should belong to exactly one concern.

---

Recommended Project Structure

Excel Workbook
│
├── ThisWorkbook
│
├── Worksheets
│   ├── Dashboard
│   ├── Configuration
│   └── Data
│
├── Standard Modules
│   ├── Constants
│   ├── Config
│   ├── Business
│   ├── Controllers
│   ├── Excel
│   ├── Logging
│   ├── Files
│   ├── Utilities
│   └── External Systems
│
├── Class Modules
│
├── UserForms
│
└── References

This organization scales much better than grouping code randomly.

---

Module Responsibilities

Every module must have one clearly defined purpose.

Examples:

Good

Utils_File
Utils_String
Utils_Array
Excel_Table
Excel_Format
Business_Orders
Business_Validation
Config
Logging

Bad

Misc
General
Helpers
VariousFunctions
Module1

The module name should immediately describe its responsibility.

---

Workbook Responsibilities

The workbook should contain only workbook lifecycle logic.

Typical responsibilities:

- Workbook_Open
- Workbook_BeforeClose
- Application initialization
- Application shutdown

Avoid implementing business logic inside ThisWorkbook.

---

Worksheet Responsibilities

Worksheet modules should only contain worksheet events.

Examples:

- SelectionChange
- Change
- Activate
- Deactivate

Business rules should never be implemented directly inside worksheet modules.

Instead, delegate to dedicated modules.

Example:

Private Sub Worksheet_Change(ByVal Target As Range)

    OrdersController.OnWorksheetChanged Target

End Sub

---

UserForm Responsibilities

UserForms should be responsible only for:

- displaying data
- collecting user input
- calling application services

UserForms should not implement business logic.

---

Business Modules

Business modules contain application logic.

Examples:

- invoice calculation
- validation
- pricing
- reporting
- workflow

Business modules should never depend directly on the user interface.

---

Utility Modules

Utility modules provide reusable generic functionality.

Examples:

- String helpers
- Array helpers
- Collection helpers
- Date helpers
- File helpers

Utility modules should remain independent from business rules.

---

Configuration Module

Application configuration should be centralized.

Never scatter configuration values across multiple modules.

Configuration examples:

- file paths
- worksheet names
- timeout values
- application settings

---

Constants Module

Application-wide constants belong in a dedicated Constants module.

Benefits:

- single source of truth
- easier maintenance
- improved readability

Avoid hardcoded values throughout the application.

---

Logging Module

Logging should be centralized.

Every module should use the same logging service.

Never mix:

- MsgBox
- Debug.Print
- custom logging

Choose one logging strategy.

---

External Systems

Interactions with external applications should be isolated.

Examples:

- SAP
- Outlook
- Word
- Access
- REST APIs
- SQL Server

Never spread COM automation across unrelated modules.

Encapsulate external communication behind dedicated modules.

---

Dependency Direction

Dependencies should flow downward.

Example:

User Interface

↓

Business

↓

Excel / SAP / File System

Infrastructure should never call business modules.

Avoid circular dependencies.

---

Low Coupling

Modules should know as little as possible about each other.

Avoid unnecessary dependencies.

Pass data through parameters whenever practical.

---

High Cohesion

Everything inside one module should be closely related.

If a module contains unrelated functionality, split it.

---

Encapsulation

Hide implementation details.

Expose only the procedures intended to be used by other modules.

Prefer:

Private Sub InternalCalculation()

Public Function CalculatePrice()

instead of exposing every procedure.

---

Public Interface

Public procedures define the module's interface.

Everything else should be Private unless there is a valid reason.

Minimize the public surface area.

---

Class Modules

Use classes whenever a concept naturally represents an object.

Examples:

- Customer
- Session
- Logger
- Configuration
- Printer
- Rule Engine

Avoid creating classes that contain only static helper functions.

---

Event Management

Event procedures should remain small.

Typical event flow:

Excel Event

↓

Validation

↓

Business Logic

↓

Update UI

Avoid placing hundreds of lines inside an event procedure.

---

State Management

Avoid unnecessary global state.

Prefer:

- local variables
- class instances
- explicit parameters

Global variables increase coupling.

---

Reusable Components

Whenever functionality may be reused:

Extract it into:

- helper module
- service
- class

Never duplicate logic.

---

Scalability

Assume the application will grow.

New features should require adding modules rather than modifying unrelated ones.

A scalable architecture minimizes regression risks.

---

Folder Thinking

Although VBA stores everything inside a workbook, mentally organize the project as if it were a modern software repository.

Think in terms of:

- packages
- namespaces
- services
- controllers
- infrastructure

This mindset naturally produces cleaner code.

---

Architectural Anti-Patterns

Avoid:

- God Modules
- God Procedures
- Circular dependencies
- Business logic in worksheets
- Business logic in UserForms
- Massive event procedures
- Random utility modules
- Hardcoded workbook references
- Copy-paste architecture

---

AI Architectural Rules

When generating a new application, the AI shall:

- Create coherent modules.
- Separate concerns.
- Minimize coupling.
- Maximize cohesion.
- Recommend classes when appropriate.
- Centralize configuration.
- Centralize constants.
- Centralize logging.
- Isolate external systems.
- Prefer composition over duplication.

---

AI Self-Review Checklist

Before presenting a solution, verify:

- Does every module have one responsibility?
- Are business rules isolated?
- Is the UI independent from the business layer?
- Is configuration centralized?
- Are constants centralized?
- Is logging centralized?
- Are dependencies logical?
- Is the architecture scalable?
- Would another developer understand the structure immediately?

If any answer is negative, improve the architecture before generating the final solution.

---

Guiding Principle

«A well-structured architecture reduces future development costs more than any individual optimization.»

A professional VBA application should be organized like any modern software project, regardless of the limitations of the VBA environment.

---

End of Chapter 02