---
name: vba-ai-skill
description: "Complete enterprise VBA standards for Excel and Office automation. Use when generating, reviewing, refactoring, or explaining VBA code in this repository, especially for architecture, naming, comments, error handling, logging, performance, Excel object model usage, templates, and code review."
user-invocable: false
---

# VBA AI Skill

Use this skill for any VBA work in this repository. The detailed annexes below are part of the skill and should be followed together, not in isolation.

## How to use this skill

- For foundational AI principles read [AI principles](./00_AI_Principles.md).
- For philosophy for building VBA applications, read [General Philosophy](./01_General_Philosophy.md).
- For architecture and organization, read [Project Architecture](./02_Project_Architecture.md) and [Module Organization](./03_Module_Organization.md).
- For style and conventions, read [Naming Convention](./04_Naming_Convention.md), [Data Types](./07_Data_Types.md), [Comments and Documentation](./08_Comments_Documentation.md), and [Code Formatting](./09_Code_Formatting.md).
- For error handling and logging, read [Error Handling](./05_Error_Handling.md) and [Logging](./10_Logging.md).
- For Excel object model usage, read [Excel Object Model](./11_Excel_Object_Model.md), [Range Best Practices](./12_Range_Best_Practices.md), and [ListObjects](./13_ListObjects.md).
- For data structures, UI flow, and events, read [Arrays, Collections, Dictionaries](./14_Arrays_Collections_Dictionaries.md), [Class Modules](./15_Class_Modules.md), [UserForms](./16_UserForms.md), and [Events](./17_Events.md).
- For external systems, read [File System](./18_File_System.md), [Windows API](./19_Windows_API.md), and [COM Automation](./20_COM_Automation.md).
- For performance, read [Performance](./21_Performance.md).
- For quality controls, read [Defensive Programming](./22_Defensive_Programming.md), [Refactoring](./23_Refactoring.md), and [Anti Patterns](./25_Anti_Patterns.md).
- For code review, read [Code Review Checklist](./24_Code_Review_Checklist.md).
- For generation rules, read [AI Generation Rules](./26_AI_Generation_Rules.md).
- For templates, read [Templates](./27_Templates.md).

## Mandatory operating rules

- Prefer production-grade VBA, not tutorial code.
- Treat Excel as an I/O layer, not the place where business logic lives.
- Keep logic explicit, deterministic, and easy to extend.
- Validate inputs early and fail fast.
- Use `Utils_Log` for diagnostics and never hide unexpected errors.
- Use the supplied templates instead of inventing new procedure shapes.
- When reviewing code, use the full checklist before considering the task complete.

## Reading order

When the task is to generate or refactor code, apply the rules in this order:

1. [core principles](./00_AI_Principles.md)
2. [foundation and philosophy](./01_General_Philosophy.md)
3. [architecture and organization](./02_Project_Architecture.md)
4. [style and conventions](./04_Naming_Convention.md)
5. [generation rules](./26_AI_Generation_Rules.md)
6. [procedures and functions](./27_Templates.md)
7. [error handling](./05_Error_Handling.md)
8. [performance](./21_Performance.md)
9. [templates](./27_Templates.md)

When the task is to review code, use the [code review checklist](./24_Code_Review_Checklist.md) as the authoritative checklist, then verify any conflicting rule against the higher-level documents above.
