VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchScraper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: MatchScraper

Public Function GetMatchResults(matchUrl As String) As MatchResult
    Dim appIE As Object
    Set appIE = CreateObject("internetexplorer.application")
    With appIE
        .Navigate matchUrl
        .Visible = False
    End With
    
    Do While appIE.Busy Or appIE.readyState <> 4
        DoEvents
    Loop
    
    'Wait for page to populate from ajax query
    Do While appIE.document.getElementById("MatchCardBody") Is Nothing
        DoEvents
    Loop
        
    Set GetMatchResults = ParseMatchCard(appIE.document)
    
    appIE.Quit
    Set appIE = Nothing
End Function

Public Function ParseMatchCard(matchCard As HTMLDocument) As MatchResult
    Dim oMatchResult As MatchResult
    Set oMatchResult = New MatchResult
    
    'TODO: Populate Match Results
    oMatchResult.Key = ""
    oMatchResult.WeekNumber = 0
    oMatchResult.Format = "9S1D"
    oMatchResult.Division = ""
    Set oMatchResult = ParseMatchSummary(oMatchResult, matchCard)
    Set oMatchResult = ParsePlayerNames(oMatchResult, matchCard)
    Set oMatchResult = ParseSinglesGameScores(oMatchResult, matchCard)
    Set oMatchResult = ParseDoublesGameScores(oMatchResult, matchCard)
    
    Set ParseMatchCard = oMatchResult
End Function

Public Function ParsePlayerNames(oMatchResult As MatchResult, matchCard As HTMLDocument) As MatchResult
    Dim cardResults As IHTMLElement
    Set cardResults = matchCard.getElementById("CardResults")
    
    oMatchResult.PlayerX = cardResults.Children(0).Children(1).innerText
    oMatchResult.PlayerY = cardResults.Children(0).Children(2).innerText
    oMatchResult.PlayerZ = cardResults.Children(0).Children(3).innerText
    oMatchResult.PlayerA = cardResults.Children(2).Children(0).innerText
    oMatchResult.PlayerB = cardResults.Children(4).Children(0).innerText
    oMatchResult.PlayerC = cardResults.Children(6).Children(0).innerText
    
    Set ParsePlayerNames = oMatchResult
End Function

Public Function ParseMatchSummary(oMatchResult As MatchResult, matchCard As HTMLDocument) As MatchResult
    Dim text As String
    Dim element As IHTMLElement
    
    Set element = matchCard.getElementsByClassName("teamNames")(0)
    oMatchResult.HomeTeam = element.Children(0).innerText
    oMatchResult.AwayTeam = element.Children(1).innerText
    
    text = matchCard.getElementsByClassName("dates")(0).innerText
    text = Trim(Split(text, ":")(1))
    oMatchResult.MatchDate = DateValue(text)
    
    text = matchCard.getElementsByClassName("matchScore")(0).innerText
    text = Trim(Split(text, ":")(1))
    oMatchResult.MatchScore = FormatScore(text)

    Set ParseMatchSummary = oMatchResult
End Function

Public Function ParseSinglesGameScores(oMatchResult As MatchResult, matchCard As HTMLDocument) As MatchResult

    Dim text As String
    Dim setNum, gameNum As Long
    Dim balance As Integer
    Dim s, G As IHTMLElement
    Dim sets, Games As IHTMLElementCollection
    
    Set sets = matchCard.getElementsByClassName("set")
    For Each s In sets
        gameNum = 0
        balance = 0
        For Each G In sets(setNum).Children
            text = FormatScore(G.innerText)
            balance = balance + CalcGameBalance(text)
            oMatchResult.Games(GetSetId(setNum)).Ends(gameNum) = text
            gameNum = gameNum + 1
        Next
        
        oMatchResult.Games(GetSetId(setNum)).Winner = CalcWinner(balance)
        setNum = setNum + 1
    Next
    
    Set ParseSinglesGameScores = oMatchResult
End Function

Public Function ParseDoublesGameScores(oMatchResult As MatchResult, matchCard As HTMLDocument) As MatchResult

    Dim text As String
    Dim gameNum As Long
    Dim balance As Integer
    Dim G As IHTMLElement
    Dim Games As IHTMLElementCollection
    
    gameNum = 0
    balance = 0
    For Each G In matchCard.getElementsByClassName("doublesSet")(0).Children
        text = FormatScore(G.innerText)
        balance = balance + CalcGameBalance(text)
        oMatchResult.Games(9).Ends(gameNum) = text
        gameNum = gameNum + 1
    Next
    
    oMatchResult.Games(9).Winner = CalcWinner(balance)
            
    Set ParseDoublesGameScores = oMatchResult
End Function

Public Function CalcGameBalance(score As String) As Integer
    Dim scores() As String
    scores() = Split(score, "~")
    
    If CInt(scores(0)) > CInt(scores(1)) Then
        CalcGameBalance = 1
    Else
        CalcGameBalance = -1
    End If

End Function

Public Function GetSetId(ByVal setNum As Integer) As Integer
    Select Case setNum
        Case 0
            GetSetId = 0
        Case 1
            GetSetId = 8
        Case 2
            GetSetId = 4
        Case 3
            GetSetId = 3
        Case 4
            GetSetId = 1
        Case 5
            GetSetId = 6
        Case 6
            GetSetId = 7
        Case 7
            GetSetId = 5
        Case 8
            GetSetId = 2
    End Select
End Function

Public Function CalcGameId(ByVal setNum As Integer, ByVal gameNum As Integer) As String
    CalcGameId = CStr(GetSetId(setNum)) + "_" + CStr(gameNum + 1)
End Function


Public Function CalcWinner(balance As Integer) As String
    If balance > 0 Then
        CalcWinner = "Home"
    Else
        CalcWinner = "Away"
    End If

End Function

Public Function FormatScore(text As String) As String
    Dim scores() As String
    scores() = Split(text, "-")
    FormatScore = Trim(scores(0)) & "~" & Trim(scores(1))
End Function


