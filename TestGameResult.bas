Attribute VB_Name = "TestGameResult"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGameResultCanBePopulatedWithData()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected, actual As GameResult
    
    Set expected = New GameResult
    expected.Winner = "Home"
    expected.Ends(0) = "11~0"
    expected.Ends(1) = "11~1"
    expected.Ends(2) = "11~1"
    expected.Ends(3) = ""
    expected.Ends(4) = ""

    'Act:
    Set actual = New GameResult
    actual.Winner = "Home"
    actual.Ends(0) = "11~0"
    actual.Ends(1) = "11~1"
    actual.Ends(2) = "11~1"
    actual.Ends(3) = ""
    actual.Ends(4) = ""
    
    'Assert:
    Assert.IsTrue expected.IsEquivalent(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
