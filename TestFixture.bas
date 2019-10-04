Attribute VB_Name = "TestFixture"
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
Private Sub TestFixtureCanBePopulatedWithData()
    On Error GoTo TestFail
    
    'Arrange:
    Dim oFixtureHelper As FixtureHelper
    Set oFixtureHelper = New FixtureHelper
    Dim expected, actual As Fixture
    
    Set expected = oFixtureHelper.CreateFixture(14, DateValue("2020-05-01"), "Test Team 1", "Test Team 2", "8~2", "https://www.tabletennis365.com/")
        
    'Act:
    Set actual = oFixtureHelper.CreateFixture(14, DateValue("2020-05-01"), "Test Team 1", "Test Team 2", "8~2", "https://www.tabletennis365.com/")
        
    'Assert:
    Assert.IsTrue expected.IsEquivalent(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

