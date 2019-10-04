Attribute VB_Name = "TestMatchHelper"
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
Private Sub CalculateMatchScoreShouldCalculateCorrectResult()
    On Error GoTo TestFail
    
    'Arrange:
    Dim data As MatchResult
    Dim expected, actual As String
    Dim service As MatchHelper
    
    Set data = New MatchResult
    
    data.Games(0).Ends(0) = "11~07"
    data.Games(0).Ends(1) = "11~07"
    data.Games(0).Ends(2) = "08~11"
    data.Games(0).Ends(3) = "11~08"
    data.Games(0).Ends(4) = ""
    data.Games(0).Winner = "Home"
    data.Games(1).Ends(0) = "11~05"
    data.Games(1).Ends(1) = "11~01"
    data.Games(1).Ends(2) = "11~08"
    data.Games(1).Ends(3) = ""
    data.Games(1).Ends(4) = ""
    data.Games(1).Winner = "Home"
    data.Games(2).Ends(0) = "08~11"
    data.Games(2).Ends(1) = "06~11"
    data.Games(2).Ends(2) = "08~11"
    data.Games(2).Ends(3) = ""
    data.Games(2).Ends(4) = ""
    data.Games(2).Winner = "Away"
    data.Games(3).Ends(0) = "11~06"
    data.Games(3).Ends(1) = "06~11"
    data.Games(3).Ends(2) = "11~04"
    data.Games(3).Ends(3) = "03~11"
    data.Games(3).Ends(4) = "11~07"
    data.Games(3).Winner = "Home"
    data.Games(4).Ends(0) = "11~08"
    data.Games(4).Ends(1) = "10~12"
    data.Games(4).Ends(2) = "06~11"
    data.Games(4).Ends(3) = "11~09"
    data.Games(4).Ends(4) = "11~04"
    data.Games(4).Winner = "Home"
    data.Games(5).Ends(0) = "11~05"
    data.Games(5).Ends(1) = "11~09"
    data.Games(5).Ends(2) = "10~12"
    data.Games(5).Ends(3) = "09~11"
    data.Games(5).Ends(4) = "11~09"
    data.Games(5).Winner = "Home"
    data.Games(6).Ends(0) = "05~11"
    data.Games(6).Ends(1) = "13~11"
    data.Games(6).Ends(2) = "11~06"
    data.Games(6).Ends(3) = "14~16"
    data.Games(6).Ends(4) = "11~08"
    data.Games(6).Winner = "Home"
    data.Games(7).Ends(0) = "08~11"
    data.Games(7).Ends(1) = "11~06"
    data.Games(7).Ends(2) = "08~11"
    data.Games(7).Ends(3) = "11~08"
    data.Games(7).Ends(4) = "11~09"
    data.Games(7).Winner = "Home"
    data.Games(8).Ends(0) = "11~07"
    data.Games(8).Ends(1) = "03~11"
    data.Games(8).Ends(2) = "12~14"
    data.Games(8).Ends(3) = "08~11"
    data.Games(8).Ends(4) = ""
    data.Games(8).Winner = "Away"
    data.Games(9).Ends(0) = "01~11"
    data.Games(9).Ends(1) = "02~11"
    data.Games(9).Ends(2) = "03~11"
    data.Games(9).Ends(3) = ""
    data.Games(9).Ends(4) = ""
    data.Games(9).Winner = "Away"
             
    expected = "7~3"
                
    'Act:
    Set service = New MatchHelper
    actual = service.CalculateMatchScore(data)
        
    'Assert:
    Assert.AreEqual expected, actual, "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub CalculateMatchScoreShouldCalculateZeroHomeResult()
    On Error GoTo TestFail
    
    'Arrange:
    Dim data As MatchResult
    Dim expected, actual As String
    Dim service As MatchHelper
    
    Set data = New MatchResult
    
    data.Games(0).Winner = "Away"
    data.Games(1).Winner = "Away"
    data.Games(2).Winner = "Away"
    data.Games(3).Winner = "Away"
    data.Games(4).Winner = "Away"
    data.Games(5).Winner = "Away"
    data.Games(6).Winner = "Away"
    data.Games(7).Winner = "Away"
    data.Games(8).Winner = "Away"
    data.Games(9).Winner = "Away"
             
    expected = "0~10"
                
    'Act:
    Set service = New MatchHelper
    actual = service.CalculateMatchScore(data)
        
    'Assert:
    Assert.AreEqual expected, actual, "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub CalculateMatchScoreShouldCalculateZeroAwayResult()
    On Error GoTo TestFail
    
    'Arrange:
    Dim data As MatchResult
    Dim expected, actual As String
    Dim service As MatchHelper
    
    Set data = New MatchResult
    
    data.Games(0).Winner = "Home"
    data.Games(1).Winner = "Home"
    data.Games(2).Winner = "Home"
    data.Games(3).Winner = "Home"
    data.Games(4).Winner = "Home"
    data.Games(5).Winner = "Home"
    data.Games(6).Winner = "Home"
    data.Games(7).Winner = "Home"
    data.Games(8).Winner = "Home"
    data.Games(9).Winner = "Home"
             
    expected = "10~0"
                
    'Act:
    Set service = New MatchHelper
    actual = service.CalculateMatchScore(data)
        
    'Assert:
    Assert.AreEqual expected, actual, "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
