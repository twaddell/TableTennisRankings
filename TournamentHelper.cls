VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TournamentHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: TournamentHelper

Public Sub CopyResults()
    Debug.Print "Copying Tournament Results"
    
    Dim tmentResults As Collection
    Set tmentResults = ReadTournamentResults()
    
    CopyToResults tmentResults
End Sub

Function ReadTournamentResults() As Collection
    Dim rg As Range
    Set rg = Sheets("Tournament Results").Range("A1").CurrentRegion
    
    Dim col As New Collection
    Dim oResult As TournamentResult
    Dim i As Long
    
    For i = 2 To rg.Rows.Count
        Set oResult = New TournamentResult
        ' Add the details
        oResult.MatchDate = rg.Cells(i, 1)
        oResult.PlayerA = rg.Cells(i, 2)
        oResult.PlayerB = rg.Cells(i, 3)
        oResult.Winner = rg.Cells(i, 4)
        col.Add oResult
    Next i
    
    Set ReadTournamentResults = col
End Function

Sub CopyToResults(tmentResults As Collection)
    Dim oMatchHelper As New MatchHelper
    Dim matchResults As New Collection
    Dim oMatchResult As MatchResult
    Dim oTmentResult As TournamentResult
    
    For Each oTmentResult In tmentResults
        Set oMatchResult = New MatchResult
        oMatchResult.Key = "99-T-TOUR-TOUR"
        oMatchResult.WeekNumber = 99
        oMatchResult.Format = "9S0D"
        oMatchResult.Division = "T"
        oMatchResult.MatchDate = oTmentResult.MatchDate
        oMatchResult.HomeTeam = "_Tournament"
        oMatchResult.AwayTeam = "_Tournament"
        If oTmentResult.Winner = "A" Then
            oMatchResult.MatchScore = "1~0"
            oMatchResult.Games(0).Winner = "Home"
        Else
            oMatchResult.MatchScore = "0~1"
            oMatchResult.Games(0).Winner = "Away"
        End If
        oMatchResult.PlayerA = oTmentResult.PlayerA
        oMatchResult.PlayerB = "Forfeit"
        oMatchResult.PlayerC = "Forfeit"
        oMatchResult.PlayerX = oTmentResult.PlayerB
        oMatchResult.PlayerY = "Forfeit"
        oMatchResult.PlayerZ = "Forfeit"
        oMatchResult.Games(0).Ends(0) = "99~99"
        oMatchResult.Games(0).Ends(1) = "99~99"
        oMatchResult.Games(0).Ends(2) = "99~99"
        oMatchResult.Games(0).Ends(3) = "99~99"
        oMatchResult.Games(0).Ends(4) = "99~99"
        
        matchResults.Add oMatchResult
    Next
    
    oMatchHelper.DeleteTournamentResults
    oMatchHelper.AppendMatchResults matchResults
    
End Sub
