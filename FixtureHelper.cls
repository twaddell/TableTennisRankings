VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixtureHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: FixtureHelper

Public Function CreateFixture(iWeekNumber As Integer, dMatchDate As Date, sHomeTeam As String, sAwayTeam As String, sMatchScore As String, sMatchCardUrl As String) As Fixture
    Dim oFixture As Fixture
    Set oFixture = New Fixture
    oFixture.WeekNumber = iWeekNumber
    oFixture.MatchDate = dMatchDate
    oFixture.HomeTeam = sHomeTeam
    oFixture.AwayTeam = sAwayTeam
    oFixture.MatchScore = sMatchScore
    oFixture.MatchCardUrl = sMatchCardUrl
    Set CreateFixture = oFixture
End Function

Public Sub AppendFixtures(fixtureSheet As Worksheet, divId As Integer, oFixtures As Collection)
    Debug.Print "Writing Fixtures for Division " + CStr(divId)
    Dim offset As Integer
    offset = (divId - 1) * 7
        
    fixtureSheet.Cells(1, 1 + offset) = "Division " + CStr(divId) + " (Rslt is a Hyperlink to Match Card)"
    With fixtureSheet.Range(fixtureSheet.Cells(1, 1 + offset), fixtureSheet.Cells(1, 6 + offset))
        .HorizontalAlignment = xlCenter
        .MergeCells = True
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Size = 12
        .Font.Underline = xlUnderlineStyleSingle
        .Font.ColorIndex = 3
    End With
    
    fixtureSheet.Cells(2, 1 + offset) = "WkNo"
    fixtureSheet.Cells(2, 2 + offset) = "Home Team"
    fixtureSheet.Cells(2, 3 + offset) = "VPR"
    fixtureSheet.Cells(2, 4 + offset) = "Away Team"
    fixtureSheet.Cells(2, 5 + offset) = "Match Date"
    fixtureSheet.Cells(2, 6 + offset) = "Rslt"
        
    Dim col As New Collection
    Dim oFixture As Fixture
    Dim i As Long
    Dim start As Long
    
    i = 3
    For Each oFixture In oFixtures
        fixtureSheet.Cells(i, 1 + offset) = oFixture.WeekNumber
        fixtureSheet.Cells(i, 2 + offset) = oFixture.HomeTeam
        fixtureSheet.Cells(i, 3 + offset) = "vs"
        fixtureSheet.Cells(i, 4 + offset) = oFixture.AwayTeam
        fixtureSheet.Cells(i, 5 + offset) = oFixture.MatchDate
        If oFixture.MatchCardUrl <> "" Then
        fixtureSheet.Hyperlinks.Add Anchor:=fixtureSheet.Cells(i, 6 + offset), Address:=oFixture.MatchCardUrl, TextToDisplay:=oFixture.MatchScore
        End If
        i = i + 1
    Next
End Sub


Public Sub ProcessFixtures(divisions As Collection)
    Dim i As Integer
    Dim divFixtures As Collection
    
    Dim oFixtureScraper As FixtureScraper
    Set oFixtureScraper = New FixtureScraper
    
    Dim fixtureSheet As Worksheet
    Set fixtureSheet = Worksheets("Fixtures")
    fixtureSheet.Cells.ClearContents
    
    For i = 1 To divisions.Count
        Set divFixtures = oFixtureScraper.GetFixtures(divisions(i))
        AppendFixtures fixtureSheet, i, divFixtures
    Next
End Sub
