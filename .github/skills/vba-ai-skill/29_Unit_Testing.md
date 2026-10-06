# 29 - Unit Testing

## Objective

Unit tests prove that small pieces of logic keep working when the code changes. They are what makes refactoring
(chapter 23) safe.

This chapter is **tool-agnostic**: the principles apply with Rubberduck's test framework, with a hand-written
test module, or with any other harness. Tool-specific details are marked as optional.

---

# Golden Rule

> A unit test checks one behaviour of one unit, quickly, without touching the outside world.

---

# 1. What to Test

Test:

- business rules (discounts, validation, label content, ZPL generation)
- parsing and formatting functions
- pure functions (chapter 06)
- classes through their **public interface** only

Do not test:

- private members (test them through the public members that call them)
- Excel, SAP or Outlook themselves
- UserForm layout (test the form model instead, chapter 16)

Code that is hard to test is usually badly structured: too many responsibilities, hidden dependencies, global state.
Fix the design (chapters 06, 15, 28) rather than skipping the test.

---

# 2. Test Structure: Arrange / Act / Assert

Every test has three clearly separated parts:

```vb
Public Sub CalculateDiscount_QuantityAboveThreshold_ReturnsTenPercent()

    ' Arrange
    Dim order As IOrder
    Set order = clsOrder.Create("A123", 11)

    ' Act
    Dim actual As Double
    actual = order.CalculateDiscount()

    ' Assert
    Assert.AreEqual 0.1, actual

End Sub
```

Rules:

- one logical assertion per test (several `Assert` lines that check one outcome are fine)
- no loops, no conditions inside a test
- no dependency between tests: each test creates what it needs and leaves nothing behind
- tests never depend on execution order

---

# 3. Naming

```
MethodUnderTest_GivenCondition_ThenExpectedResult
```

or, when the method is obvious:

```
Condition_ExpectedResult
```

Examples:

- `TryParseQuantity_EmptyText_ReturnsFalse`
- `Create_NegativeQuantity_RaisesInvalidArgument`
- `IsValid_NoOrderId_ReturnsFalse`

A failing test name must tell the reader what broke without opening the test.

(This is the one place where underscores in a procedure name are expected, next to event handlers.)

---

# 4. Tests Must Be Isolated and Fast

A unit test must not:

- read or write a file
- open a workbook
- talk to SAP, Outlook, a printer or a database
- show a `MsgBox` or any dialog (it blocks automated runs)
- depend on the current date, the active sheet, the user, the machine

Isolate those through interfaces and injection (chapter 28), and replace them with test doubles.

---

# 5. Stubs and Mocks

A **test double** replaces a real dependency.

| Type | Purpose |
|------|---------|
| Stub | Returns canned answers so the unit under test can run |
| Mock | Also records calls, so the test can verify that an interaction happened |

A stub of `IFileProvider`:

```vb
' clsStubFileProvider
Option Explicit

Implements IFileProvider

Private m_content As String

Public Property Let Content(ByVal value As String)
    m_content = value
End Property

Private Function IFileProvider_FileExists(ByVal path As String) As Boolean
    IFileProvider_FileExists = True
End Function

Private Function IFileProvider_ReadAllText(ByVal path As String) As String
    IFileProvider_ReadAllText = m_content
End Function
```

A mock adds counters or captured arguments (`SendCount`, `LastRecipient`) that the test reads in its Assert section.

Use one stub or mock per interface and keep them in the test project, never in production modules.

---

# 6. Testing Error Paths

Verify that invalid input raises the expected error, using the project error enum (chapter 05):

```vb
Public Sub Create_EmptyOrderId_RaisesEmptyArgument()

    On Error Resume Next
    Dim order As IOrder
    Set order = clsOrder.Create(vbNullString, 1)
    Dim actualError As Long
    actualError = Err.Number
    On Error GoTo 0

    Assert.AreEqual AppError.ErrEmptyArgument, actualError

End Sub
```

This is one of the rare acceptable uses of `On Error Resume Next`: it is confined to a test, around one call.
Rubberduck offers `'@ExpectedError` for the same purpose.

---

# 7. Testing Procedures with Side Effects on Excel

Separate the logic from the Excel access:

1. read the data from the worksheet into an array or a model object (thin, not unit-tested)
2. process it in a function that takes and returns arrays or objects (unit-tested)
3. write the result back (thin, not unit-tested)

The middle step carries the rules, and it is tested without any worksheet.

Integration tests that do use a real workbook are valid, but they are separate, slower, and clearly named
(`Integration_...`), and they never run as part of the quick unit suite.

---

# 8. Test Organisation

- one test module per class or module under test: `Tests_clsOrder`, `Tests_Utils_Guard`
- test modules live in a test project or a clearly separated folder, never in the delivered workbook
- shared setup goes in a dedicated helper, not copied between tests
- keep tests as readable as production code: same formatting and naming rules apply

---

# 9. Optional: Rubberduck

If Rubberduck is installed, its test framework provides `'@TestModule`, `'@TestMethod`,
`Assert.AreEqual`, `Assert.IsTrue`, `Assert.Fail`, `'@ModuleInitialize`, `'@TestInitialize`, `'@TestCleanup`
and a Test Explorer to run and group tests.

Do not assume Rubberduck is available. When it is not, write plain tests that call `Debug.Print` or raise
an error on failure, following the same Arrange / Act / Assert structure, and keep the test names portable.

---

# 10. Cyclomatic Complexity and Test Count

The number of independent execution paths of a procedure is the minimum number of tests it needs.
This is another reason to keep procedures at low complexity (chapter 06): complexity above ~5 means too many
tests and too many hidden behaviours.

---

# 11. AI Rules

When generating code, AI should:

- propose unit tests for business rules and parsing/formatting functions
- structure tests as Arrange / Act / Assert and name them `Method_Condition_Result`
- test only the public interface
- keep tests independent, fast and free of I/O, dialogs and environment dependencies
- use stubs and mocks implementing the interfaces of chapter 28
- test the error path of every guard (expected error number)
- never put test code in the delivered production modules
- say clearly when a test framework (Rubberduck) is assumed

---

# Golden Rules

1. One test, one behaviour.
2. Arrange / Act / Assert.
3. Names describe the scenario and the expectation.
4. Test the public interface only.
5. Isolate I/O behind interfaces; use stubs and mocks.
6. Tests are independent, fast and deterministic.
7. Test error paths as well as the nominal path.
8. Hard to test means badly designed.
