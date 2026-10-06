# 28 - Interfaces and Dependency Injection

## Objective

VBA has no inheritance, but it has **interfaces** (`Implements`). Combined with dependency injection and factory
methods, they make code decoupled, replaceable and testable.

Use this chapter when a class touches something that must be replaceable: the file system, a SAP session,
Outlook, a database, the clock, a printer.

This chapter builds on chapter 15 (class modules) and prepares chapter 29 (unit testing).

---

# Golden Rule

> Depend on abstractions (interfaces), not on concrete implementations. Create objects in one place, use them everywhere else.

---

# 1. Interfaces with `Implements`

An interface is a class module that declares members and contains **no logic**.

```vb
' IFileProvider
Option Explicit

Public Function FileExists(ByVal path As String) As Boolean
End Function

Public Function ReadAllText(ByVal path As String) As String
End Function
```

A class implements it:

```vb
' clsFileSystemProvider
Option Explicit

Implements IFileProvider

Private Function IFileProvider_FileExists(ByVal path As String) As Boolean
    IFileProvider_FileExists = Len(Dir$(path)) > 0
End Function

Private Function IFileProvider_ReadAllText(ByVal path As String) As String
    ' real implementation
End Function
```

Rules:

- interface names start with `I` and have no `cls` prefix (`IFileProvider`, `ICommand`, `ILogger`)
- implemented members are `Private` and named `Interface_Member`; this is the one legitimate use of an underscore
  in a procedure name besides event handlers (chapter 04)
- an interface contains no code, no state, no event
- the consumer declares a variable of the **interface type** and never of the implementation type

---

# 2. Interface Segregation

Keep interfaces small and focused: a few members that belong together.

Bad:

```vb
' IEverything: ReadFile, WriteFile, SendMail, PrintLabel, QuerySap
```

Good:

```vb
' IFileProvider, IMailSender, ILabelPrinter, ISapSession
```

A class that needs one capability should not be forced to depend on ten.

---

# 3. Composition over Inheritance

VBA does not support implementation inheritance. Do not simulate it by copying code or by chaining
default instances. Compose objects instead: a class holds references to collaborators (typed as interfaces)
and delegates to them.

```vb
' clsOrderImporter holds an IFileProvider and an ILogger, it does not derive from anything
```

---

# 4. Dependency Injection

A class must not create its own collaborators with `New` or `CreateObject`: it receives them.

Bad (hidden dependency, untestable):

```vb
Public Function LoadOrders(ByVal path As String) As Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' ...
End Function
```

Good:

```vb
Private m_files As IFileProvider

Public Sub Init(ByVal files As IFileProvider)
    Utils_Guard.NotNothing files, "files"
    Set m_files = files
End Sub
```

## Property injection

For a dependency needed for the whole life of the object:

```vb
Public Property Set FileProvider(ByVal value As IFileProvider)
    Set m_files = value
End Property
```

## Method injection

For a dependency needed by a single operation, pass it as a parameter:

```vb
Public Sub Import(ByVal files As IFileProvider, ByVal path As String)
```

## Constructor-like injection

Prefer a factory method (section 5) that receives the dependencies and returns a fully initialised object.

---

# 5. Factory Methods (`Create`)

VBA classes cannot have parameterised constructors. A **factory method** called `Create` fills that role.

It is defined on the class itself and called through its **default instance** (a global object named after the
class, available when the module's hidden attribute `VB_PredeclaredId` is `True`).

```vb
' clsOrder
Option Explicit

Private Type TState
    OrderId As String
    Quantity As Long
End Type

Private This As TState

Public Function Create(ByVal orderId As String, ByVal quantity As Long) As IOrder

    Utils_Guard.NotEmpty orderId, "orderId"
    Utils_Guard.NotNegative quantity, "quantity"

    Dim result As clsOrder
    Set result = New clsOrder

    result.OrderId = orderId
    result.Quantity = quantity

    Set Create = result

End Function

Public Property Get OrderId() As String
    OrderId = This.OrderId
End Property

Friend Property Let OrderId(ByVal value As String)
    This.OrderId = value
End Property
```

Usage:

```vb
Dim order As IOrder
Set order = clsOrder.Create("A123", 10)
```

Key points:

- `Create` returns the **interface**: callers see read-only properties; the writable `Friend` members stay hidden.
  The returned object is effectively immutable and always valid.
- Writable members that only the factory needs are `Friend`, not `Public`.
- The **default instance must stay stateless** (chapter 15): `Create` uses no instance field of the default instance.
  Guard instance members that are not meant to be called on it.
- `VB_PredeclaredId` is a hidden attribute: it is set by exporting the module, editing the attribute line
  (`Attribute VB_PredeclaredId = True`) and re-importing it, or with a tool such as Rubberduck (`'@PredeclaredId`).
  Document that requirement in the module header and tell the user when generating such a class.
- When `VB_PredeclaredId` cannot be set, use the simple `Init` approach of chapter 15.

---

# 6. Abstract Factory

When the choice of implementation depends on a condition (production vs test, SAP vs file source), wrap the choice
in a factory that itself implements an interface:

```vb
' IProviderFactory: Public Function CreateFileProvider() As IFileProvider
' clsProductionFactory: returns clsFileSystemProvider
' clsTestFactory:       returns a stub
```

Consumers receive an `IProviderFactory` and never know which implementation they get.

---

# 7. Composition Root

Objects are **created in one place**: the entry point of the application, as close as possible to where it starts
(`Workbook_Open`, the macro entry point, or a dedicated `Composition` module). Everything below receives
its collaborators.

```vb
Public Sub RunImport()

    On Error GoTo CleanFail

    Dim files As IFileProvider
    Set files = New clsFileSystemProvider

    Dim importer As clsOrderImporter
    Set importer = New clsOrderImporter
    importer.Init files

    importer.Import ThisWorkbook.Path & "\orders.csv"

CleanExit:
    Exit Sub

CleanFail:
    Utils_Log.Error Err, "RunImport"
    Resume CleanExit

End Sub
```

Rules:

- `New` appears in the composition root and in factory methods, almost nowhere else
- no service locator or global container: a global registry hides dependencies the same way global variables do
- the composition root is the only place that knows concrete types

---

# 8. What to Abstract

Abstract anything that reaches outside the VBA process or depends on the environment:

| Dependency | Interface idea |
|-----------|----------------|
| File system, dialogs | `IFileProvider`, `IFilePicker` |
| SAP GUI session | `ISapSession` |
| Outlook, Word | `IMailSender`, `IDocumentWriter` |
| Printers | `ILabelPrinter` |
| Date and time | `IClock` |
| Logging | `ILogger` |
| User prompts (`MsgBox`, `InputBox`) | `IPrompt` (`MsgBox` blocks automated tests) |

Do **not** abstract pure computation or simple data holders: an interface per class is over-engineering.
An interface is justified when there are, or soon will be, **two implementations** (real and test), or when
it isolates an unstable external system.

---

# 9. Commands (Optional Pattern)

An `ICommand` interface (`Execute`, optionally `CanExecute`) lets a UserForm button or a worksheet event trigger
an action without knowing its implementation, and keeps event handlers thin (chapter 16 and 17).

---

# 10. Error Handling

Interface members follow the procedure roles of chapter 05: guard the arguments, let plain members propagate,
and clean up in `CleanExit` when a member owns a resource.

An implementation must honour the contract of the interface (same preconditions, same error behaviour).
Document the errors an interface member may raise next to its declaration.

---

# 11. AI Rules

When generating interfaces and injected code, AI must:

- declare dependencies as interface types, never as concrete classes
- never call `New` / `CreateObject` for an external system inside a business class
- keep interfaces small and free of logic
- name interfaces with an `I` prefix, implementing classes with `cls`
- name implemented members `Interface_Member` and mark them `Private`
- return the interface from `Create` factory methods and keep the default instance stateless
- say when a class needs the hidden `VB_PredeclaredId` attribute
- create objects in the composition root
- not introduce an interface for a class that has a single implementation and no external dependency

---

# Golden Rules

1. Program to interfaces, not implementations.
2. Interfaces are small and contain no logic.
3. Inject dependencies; never create them inside the class.
4. Create objects in the composition root.
5. `Create` factories return interfaces; the default instance stays stateless.
6. Compose; do not simulate inheritance.
7. Abstract I/O and external systems, nothing more.
