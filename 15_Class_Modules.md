# 15 - Class Modules

## Objective

Class Modules are the **highest level of abstraction in VBA** and are essential for building maintainable, scalable, and testable code.

They allow VBA to move from procedural scripting to **object-oriented design (lightweight OOP)**.

This chapter defines strict rules for designing, using, and structuring class modules.

---

# Golden Rule

> Use Class Modules to model business entities and responsibilities — not to store procedural logic.

---

# 1. When to Use Class Modules

Use classes when you need:

- encapsulation of data + behavior
- reusable business logic
- stateful objects
- abstraction over Excel structure
- separation of concerns

---

## Good use cases

- `Customer`
- `Order`
- `Invoice`
- `ZplLabelBuilder`
- `ExcelTableReader`
- `LoggerWrapper`

---

## Bad use cases

- utility functions only
- random helper procedures
- single-use scripts
- macro entry points

---

# 2. Class Design Principle

> A class must represent ONE responsibility.

If a class does multiple things → split it.

---

# 3. Class Structure Rules

A clean class must follow this order:

1. Private variables
2. Property Let/Get
3. Public methods
4. Private helper methods

---

## Example structure

```vb
Option Explicit

' =========================
' Private Fields
' =========================
Private pOrderId As String
Private pQuantity As Long

' =========================
' Properties
' =========================
Public Property Get OrderId() As String
    OrderId = pOrderId
End Property

Public Property Let OrderId(ByVal value As String)
    pOrderId = value
End Property

' =========================
' Public Methods
' =========================
Public Function IsValid() As Boolean
    IsValid = (pOrderId <> "")
End Function

' =========================
' Private Methods
' =========================
Private Function NormalizeId(ByVal value As String) As String
    NormalizeId = UCase(value)
End Function
```

---

# 4. Encapsulation Rule

> All fields must be Private.

No public variables are allowed inside classes.

Bad:

```vb
Public OrderId As String
```

Good:

```vb
Private pOrderId As String
```

---

# 5. Property Naming Convention

Private fields must use prefix:

```
p + PropertyName
```

Example:

- `pOrderId`
- `pCustomerName`
- `pTotalAmount`

---

# 6. Class Instantiation Rule

Always explicitly instantiate objects:

```vb
Dim order As Order
Set order = New Order
```

Never rely on implicit creation.

---

# 7. Constructor Pattern (Init Method)

VBA does not support constructors properly.

Use explicit initialization:

```vb
Public Sub Init(ByVal orderId As String, ByVal qty As Long)
    pOrderId = orderId
    pQuantity = qty
End Sub
```

Usage:

```vb
Dim order As Order
Set order = New Order
order.Init "A123", 10
```

---

# 8. Avoid God Classes

Bad design:

- `UtilsManager`
- `GlobalHandler`
- `ExcelProcessor`

These violate single responsibility.

Split responsibilities:

- `OrderBuilder`
- `OrderValidator`
- `OrderExporter`

---

# 9. Business Logic Placement Rule

> Business logic must live inside classes, not in modules.

Bad:

```vb
If qty > 10 Then
    discount = 0.1
End If
```

Good:

```vb
discount = order.CalculateDiscount()
```

---

# 10. Class State Management

Classes may hold state:

```vb
Private pData As Variant
```

Rules:

- state must be controlled
- avoid uncontrolled global state
- reset state explicitly if reusable

---

# 11. Dependency Injection (Lightweight)

Pass dependencies via Init:

```vb
Public Sub Init(ByVal ws As Worksheet)
    Set pSheet = ws
End Sub
```

Avoid hardcoding dependencies inside class.

---

# 12. Avoid Active Dependencies

Never use inside classes:

- ActiveSheet
- ActiveWorkbook
- Selection

Classes must be context-independent.

---

# 13. Error Handling in Classes

Each class must handle its own errors OR propagate them clearly.

Preferred pattern:

```vb
On Error GoTo ErrorHandler

' logic here

Exit Function

ErrorHandler:
    Utils_Log.Error Err, "Order.CalculateDiscount"
    Err.Raise Err.Number, Err.Source, Err.Description
```

---

# 14. Class vs Module Decision

| Case | Use |
|------|-----|
| Stateless utility | Module |
| Business entity | Class |
| Workflow orchestration | Module |
| Data structure with behavior | Class |
| Excel interaction wrapper | Class |

---

# 15. Class Naming Rules

Classes must be:

- singular
- business-oriented
- clear responsibility

Good:

- `Order`
- `Invoice`
- `ZplLabel`
- `ExcelTable`

Bad:

- `OrdersManager`
- `DataUtils`
- `HelperClass`

---

# 16. Anti-Patterns

## 16.1 Property-free classes

Bad:

```vb
Public Sub DoSomething()
```

(no state, no structure)

---

## 16.2 Mixed responsibilities

Bad:

- Excel access + business logic + logging in same class

---

## 16.3 Public fields

Always forbidden.

---

## 16.4 Hidden dependencies

Bad:

```vb
Set ws = ActiveSheet
```

---

# 17. Testing Advantage

Classes allow:

- isolated testing
- predictable behavior
- reusable components

Even in VBA, this improves reliability significantly.

---

# 18. Logging inside classes

Always use structured logging:

```vb
Utils_Log.Debug "Order", "Calculating discount for " & pOrderId
```

---

# 19. AI Rules

When generating class modules, AI must:

- always use private fields
- always follow Property pattern
- never use public variables
- avoid ActiveSheet / ActiveWorkbook
- use Init method for setup
- ensure single responsibility
- separate business logic into methods
- avoid god classes
- use meaningful business names
- integrate Utils_Log for traceability

---

# Example Class

```vb
Option Explicit

Private pOrderId As String
Private pQuantity As Long

Public Sub Init(ByVal orderId As String, ByVal qty As Long)
    pOrderId = orderId
    pQuantity = qty
End Sub

Public Function IsValid() As Boolean
    IsValid = (pOrderId <> "" And pQuantity > 0)
End Function

Public Function GetLabelText() As String
    GetLabelText = "Order: " & pOrderId & " Qty: " & pQuantity
End Function
```

---

# Golden Rules

1. One class = one responsibility.
2. All fields must be private.
3. No public variables.
4. No ActiveSheet inside classes.
5. Use Init for construction.
6. Business logic belongs in classes.
7. Avoid god classes.
8. Classes must be reusable and testable.
9. Use structured logging inside methods.
10. Design classes as business objects, not helpers.