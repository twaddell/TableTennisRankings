VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: MatchHelper

Public Sub DeleteTournamentResults()
    Set rg = Sheets("Results").Range("A1").CurrentRegion
    Dim i As Long
    i = rg.Rows.Count
    Do While i >= 1
        If rg.Cells(i, 4) = "T" Then
                rg.Rows(i).EntireRow.Delete
        End If
        i = i - 1
    Loop
End Sub

Public Sub AppendMatchResults(results As Collection)
    Debug.Print "Writing Match Results"
    
    Set rg = Sheets("Results").Range("A1").CurrentRegion
    
    Dim col As New Collection
    Dim oResult As MatchResult
    Dim i As Long
    Dim start As Long
    
    i = rg.Rows.Count + 1
    For Each oResult In results
        UpdateMatchResult oResult, i
        i = i + 1
    Next
End Sub

Public Function CalculateMatchScore(oMatchResult As MatchResult) As String

Dim HomeScore, AwayScore, i As Integer
Dim oGame As GameResult

HomeScore = 0
AwayScore = 0

For i = 0 To 9
    Set oGame = oMatchResult.Games(i)
    If oGame.Winner = "Home" Then
        HomeScore = HomeScore + 1
    ElseIf oGame.Winner = "Away" Then
        AwayScore = AwayScore + 1
    End If
Next i

CalculateMatchScore = "" & HomeScore & "~" & AwayScore

End Function
Public Function ValidateMatch(ExpectedMatchScore As String, oMatchResult As MatchResult) As Boolean
    Dim isValid As Boolean
    isValid = ExpectedMatchScore = oMatchResult.MatchScore
    isValid = isValid And oMatchResult.MatchScore = CalculateMatchScore(oMatchResult)
        
    ValidateMatch = isValid
End Function

Public Sub UpdateMatchResult(oResult As MatchResult, ByVal rowIndex As Integer)

    Set rg = Sheets("Results").Range("A1").CurrentRegion
    
    If oResult.Key <> "99-T-TOUR-TOUR" Then
        If DateValue(rg.Cells(rowIndex, 5)) <> DateValue(oResult.MatchDate) Then
            rg.Cells(rowIndex, 5).Interior.ColorIndex = 4
        End If
        If rg.Cells(rowIndex, 8) <> oResult.MatchScore Then
            rg.Cells(rowIndex, 8).Interior.ColorIndex = 4
        End If
        If rg.Cells(rowIndex, 9) <> oResult.PlayerA Then
            rg.Cells(rowIndex, 9).Interior.ColorIndex = 4
        End If
        If rg.Cells(rowIndex, 10) <> oResult.PlayerB Then
            rg.Cells(rowIndex, 10).Interior.ColorIndex = 4
        End If
        If rg.Cells(rowIndex, 11) <> oResult.PlayerC Then
            rg.Cells(rowIndex, 11).Interior.ColorIndex = 4
        End If
        If rg.Cells(rowIndex, 12) <> oResult.PlayerX Then
            rg.Cells(rowIndex, 12).Interior.ColorIndex = 4
        End If
        If rg.Cells(rowIndex, 13) <> oResult.PlayerY Then
            rg.Cells(rowIndex, 13).Interior.ColorIndex = 4
        End If
        If rg.Cells(rowIndex, 14) <> oResult.PlayerZ Then
            rg.Cells(rowIndex, 14).Interior.ColorIndex = 4
        End If
    End If
    
    rg.Cells(rowIndex, 1) = oResult.Key
    'rg.Cells(rowIndex, 2) = oResult.WeekNumber
    rg.Cells(rowIndex, 3) = oResult.Format
    rg.Cells(rowIndex, 4) = oResult.Division
    rg.Cells(rowIndex, 5) = oResult.MatchDate
    rg.Cells(rowIndex, 6) = oResult.HomeTeam
    rg.Cells(rowIndex, 7) = oResult.AwayTeam
    rg.Cells(rowIndex, 8) = oResult.MatchScore
    rg.Cells(rowIndex, 9) = oResult.PlayerA
    rg.Cells(rowIndex, 10) = oResult.PlayerB
    rg.Cells(rowIndex, 11) = oResult.PlayerC
    rg.Cells(rowIndex, 12) = oResult.PlayerX
    rg.Cells(rowIndex, 13) = oResult.PlayerY
    rg.Cells(rowIndex, 14) = oResult.PlayerZ
    rg.Cells(rowIndex, 15) = oResult.Games(0).Ends(0)
    rg.Cells(rowIndex, 16) = oResult.Games(0).Ends(1)
    rg.Cells(rowIndex, 17) = oResult.Games(0).Ends(2)
    rg.Cells(rowIndex, 18) = oResult.Games(0).Ends(3)
    rg.Cells(rowIndex, 19) = oResult.Games(0).Ends(4)
    rg.Cells(rowIndex, 20) = oResult.Games(0).Winner
    rg.Cells(rowIndex, 21) = oResult.Games(1).Ends(0)
    rg.Cells(rowIndex, 22) = oResult.Games(1).Ends(1)
    rg.Cells(rowIndex, 23) = oResult.Games(1).Ends(2)
    rg.Cells(rowIndex, 24) = oResult.Games(1).Ends(3)
    rg.Cells(rowIndex, 25) = oResult.Games(1).Ends(4)
    rg.Cells(rowIndex, 26) = oResult.Games(1).Winner
    rg.Cells(rowIndex, 27) = oResult.Games(2).Ends(0)
    rg.Cells(rowIndex, 28) = oResult.Games(2).Ends(1)
    rg.Cells(rowIndex, 29) = oResult.Games(2).Ends(2)
    rg.Cells(rowIndex, 30) = oResult.Games(2).Ends(3)
    rg.Cells(rowIndex, 31) = oResult.Games(2).Ends(4)
    rg.Cells(rowIndex, 32) = oResult.Games(2).Winner
    rg.Cells(rowIndex, 33) = oResult.Games(3).Ends(0)
    rg.Cells(rowIndex, 34) = oResult.Games(3).Ends(1)
    rg.Cells(rowIndex, 35) = oResult.Games(3).Ends(2)
    rg.Cells(rowIndex, 36) = oResult.Games(3).Ends(3)
    rg.Cells(rowIndex, 37) = oResult.Games(3).Ends(4)
    rg.Cells(rowIndex, 38) = oResult.Games(3).Winner
    rg.Cells(rowIndex, 39) = oResult.Games(4).Ends(0)
    rg.Cells(rowIndex, 40) = oResult.Games(4).Ends(1)
    rg.Cells(rowIndex, 41) = oResult.Games(4).Ends(2)
    rg.Cells(rowIndex, 42) = oResult.Games(4).Ends(3)
    rg.Cells(rowIndex, 43) = oResult.Games(4).Ends(4)
    rg.Cells(rowIndex, 44) = oResult.Games(4).Winner
    rg.Cells(rowIndex, 45) = oResult.Games(5).Ends(0)
    rg.Cells(rowIndex, 46) = oResult.Games(5).Ends(1)
    rg.Cells(rowIndex, 47) = oResult.Games(5).Ends(2)
    rg.Cells(rowIndex, 48) = oResult.Games(5).Ends(3)
    rg.Cells(rowIndex, 49) = oResult.Games(5).Ends(4)
    rg.Cells(rowIndex, 50) = oResult.Games(5).Winner
    rg.Cells(rowIndex, 51) = oResult.Games(6).Ends(0)
    rg.Cells(rowIndex, 52) = oResult.Games(6).Ends(1)
    rg.Cells(rowIndex, 53) = oResult.Games(6).Ends(2)
    rg.Cells(rowIndex, 54) = oResult.Games(6).Ends(3)
    rg.Cells(rowIndex, 55) = oResult.Games(6).Ends(4)
    rg.Cells(rowIndex, 56) = oResult.Games(6).Winner
    rg.Cells(rowIndex, 57) = oResult.Games(7).Ends(0)
    rg.Cells(rowIndex, 58) = oResult.Games(7).Ends(1)
    rg.Cells(rowIndex, 59) = oResult.Games(7).Ends(2)
    rg.Cells(rowIndex, 60) = oResult.Games(7).Ends(3)
    rg.Cells(rowIndex, 61) = oResult.Games(7).Ends(4)
    rg.Cells(rowIndex, 62) = oResult.Games(7).Winner
    rg.Cells(rowIndex, 63) = oResult.Games(8).Ends(0)
    rg.Cells(rowIndex, 64) = oResult.Games(8).Ends(1)
    rg.Cells(rowIndex, 65) = oResult.Games(8).Ends(2)
    rg.Cells(rowIndex, 66) = oResult.Games(8).Ends(3)
    rg.Cells(rowIndex, 67) = oResult.Games(8).Ends(4)
    rg.Cells(rowIndex, 68) = oResult.Games(8).Winner
    rg.Cells(rowIndex, 69) = oResult.Games(9).Ends(0)
    rg.Cells(rowIndex, 70) = oResult.Games(9).Ends(1)
    rg.Cells(rowIndex, 71) = oResult.Games(9).Ends(2)
    rg.Cells(rowIndex, 72) = oResult.Games(9).Ends(3)
    rg.Cells(rowIndex, 73) = oResult.Games(9).Ends(4)
    rg.Cells(rowIndex, 74) = oResult.Games(9).Winner

End Sub
