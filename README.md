# VBA AI Skill

> **Enterprise-grade AI Coding Standards for Microsoft Excel VBA**

Version: 1.0.0
Author: Daniel Moindrot
Language: English


# Overview

**VBA AI Skill** is a comprehensive coding standard designed specifically for AI coding assistants that generate Microsoft Excel VBA code.

Unlike traditional style guides intended for human developers, this project defines **strict behavioral rules** that an AI must follow when designing, generating, reviewing and refactoring VBA applications.

The objective is to ensure that AI-generated code is indistinguishable from code written by an experienced software engineer.

The standards are technology-agnostic for VBA and can be combined with project-specific contexts such as SAP GUI Scripting, ERP automation, reporting tools or custom Excel applications.


# Philosophy

This project is based on one simple principle:

> **An AI should generate production-quality code by default.**

Generated code should never look like an example found in a tutorial.

Instead, every generated module should be suitable for direct integration into a professional application.


# Objectives

The AI must always generate code that is:

* Correct
* Readable
* Maintainable
* Reusable
* Predictable
* Robust
* Well documented
* Efficient
* Easy to review
* Easy to extend

The goal is not to generate the shortest code.

The goal is to generate the best long-term solution.


# Target Audience

This repository is intended for:

* ChatGPT Projects
* Claude Projects
* Cursor Rules
* GitHub Copilot Instructions
* Windsurf
* Continue.dev
* Roo Code
* Cline
* Any AI coding assistant supporting custom instructions

It can also serve as an internal VBA development standard for software teams.


# Repository Structure

Files 01 to 27 should be created soon.

```text
VBA-AI-Skill/
│
├── README.md
│
├── 00_AI_Principles.md
├── 01_General_Philosophy.md 
├── 02_Project_Architecture.md
├── 03_Module_Organization.md
├── 04_Naming_Conventions.md
├── 05_Code_Formatting.md
├── 06_Comments_Documentation.md
├── 07_Data_Types.md
├── 08_Functions_and_Procedures.md
├── 09_Error_Handling.md
├── 10_Logging.md
├── 11_Excel_Object_Model.md
├── 12_Range_Best_Practices.md
├── 13_ListObjects.md
├── 14_Arrays_Collections_Dictionaries.md
├── 15_Class_Modules.md
├── 16_UserForms.md
├── 17_Events.md
├── 18_File_System.md
├── 19_Windows_API.md
├── 20_COM_Automation.md
├── 21_Performance.md
├── 22_Defensive_Programming.md
├── 23_Refactoring.md
├── 24_Code_Review_Checklist.md
├── 25_Anti_Patterns.md
├── 26_AI_Generation_Rules.md
├── 27_Templates.md
│
└── CHANGELOG.md
```


# Design Philosophy

The VBA AI Skill follows modern software engineering principles adapted specifically for Microsoft Office VBA.

The primary design goals are:

* Clean Code
* SOLID principles (adapted to VBA)
* DRY
* KISS
* YAGNI
* Defensive Programming
* Explicit Programming
* Production-first mindset


# AI Responsibilities

Whenever generating VBA code, the AI must:

* Write complete solutions.
* Produce production-ready code.
* Avoid obsolete VBA practices.
* Follow all naming conventions.
* Use structured error handling.
* Release object references correctly.
* Prefer reusable components.
* Explain important design decisions.
* Recommend refactoring when appropriate.

The AI should behave like a senior software engineer, not like a code completion tool.


# Scope

This standard covers:

* Excel VBA
* Office COM automation
* Excel Object Model
* Class Modules
* UserForms
* Collections
* Dictionaries
* Arrays
* File System
* Windows API
* COM programming
* Performance optimization
* Defensive programming
* Error handling
* Logging
* Refactoring
* Documentation
* Code review

Project-specific frameworks (SAP, Outlook, Word, Access, etc.) should define additional rules on top of this standard.


# Rule Hierarchy

When multiple rules apply, they shall be interpreted in the following order:

1. Project-specific rules
2. VBA AI Skill rules
3. Microsoft Office VBA recommendations
4. General software engineering best practices

Project-specific rules always override generic VBA rules.


# Guiding Principles

Every generated solution should satisfy the following principles:

* Simplicity
* Consistency
* Explicitness
* Maintainability
* Predictability
* Reusability
* Reliability

If two solutions solve the same problem, always choose the one that is easier to understand.


# What This Repository Is Not

This repository is **not**:

* a VBA tutorial
* a language reference
* a beginner guide
* an Excel handbook

Instead, it defines **how an AI should think while writing VBA**.


# Coding Standards

Every generated solution must follow strict conventions regarding:

* module organization
* naming
* formatting
* documentation
* error handling
* logging
* object lifecycle
* performance
* architecture

These standards are described in the following chapters.


# Continuous Improvement

This repository is intended to evolve.

New standards should be added when:

* recurring code smells are identified
* better design patterns emerge
* Microsoft introduces new VBA capabilities
* AI-generated code reveals recurring weaknesses

Standards should never be removed without strong justification.


# Contributing Principles

Every new rule should satisfy the following requirements:

* Improve code quality.
* Reduce ambiguity.
* Increase maintainability.
* Be technology independent whenever possible.
* Include practical examples.
* Avoid subjective preferences unless they provide measurable benefits.


# Vision

The long-term objective of this project is to become a complete reference for AI-assisted VBA development.

When an AI is instructed to use this skill, the generated code should consistently exhibit the quality expected from an experienced enterprise software engineer.

The ideal outcome is simple:

> The user should never have to ask the AI to "clean up", "refactor", or "make the code more professional" after the initial generation.


# License

This project is intended as a reusable coding standard and may be adapted to any VBA-based project.

Project-specific conventions should be implemented as separate extension documents that build upon this core standard.


**"Good code solves today's problem. Great code is still easy to understand five years later."**
