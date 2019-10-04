VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixtureScraper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: FixtureScraper
Public Function GetFixtures(fixturesUrl As String) As Collection
    Debug.Print "Gettings Fixtures from " + fixturesUrl
    Dim appIE As Object
    Set appIE = CreateObject("internetexplorer.application")
    With appIE
        .Navigate fixturesUrl
        .Visible = False
    End With
    
    Do While appIE.Busy Or appIE.readyState <> 4
        DoEvents
    Loop
    
    'Wait for page to populate from ajax query
    Do While appIE.document.getElementById("Fixtures") Is Nothing
        DoEvents
    Loop
        
    Set GetFixtures = ParseFixtures(appIE.document)
    
    appIE.Quit
    Set appIE = Nothing
End Function

Public Function ParseFixtures(fixturesList As HTMLDocument) As Collection
    Dim oFixtures As Collection
    Dim weeks As IHTMLElementCollection4
    Dim week As IHTMLDivElement
    Dim WeekNumber As Integer
            
    Set weeks = fixturesList.getElementsByClassName("fixtureWeek")
    Set oFixtures = New Collection
    For Each week In weeks
        WeekNumber = ParseWeekNumber(week.Children(0).innerText)
        Set oFixtures = ParseWeekFixtures(oFixtures, WeekNumber, week.Children(1))
    Next
        
    Set ParseFixtures = oFixtures
End Function

Public Function ParseWeekFixtures(oFixtures As Collection, WeekNumber As Integer, weekFixtures As IHTMLElement) As Collection
    Dim weekFixture As IHTMLElement
    Dim doc As HTMLDocument
    
    For Each weekFixture In weekFixtures.Children
        Dim oFixure As Fixture
        
        If (weekFixture.className = "fixture" Or InStr(weekFixture.className, "fixture ") > 0) Then
            Set oFixture = New Fixture
                        
            Set doc = New HTMLDocument
            doc.body.innerHTML = weekFixture.innerHTML
                    
            oFixture.WeekNumber = WeekNumber
            oFixture.MatchDate = DateValue(weekFixture.Children(2).Children(0).getAttribute("datetime"))
            oFixture.HomeTeam = doc.getElementsByClassName("homeTeam")(0).Children(0).innerText
            oFixture.AwayTeam = doc.getElementsByClassName("awayTeam")(0).Children(0).innerText
            
            Dim scores As IHTMLElementCollection
            Set scores = doc.getElementsByClassName("score")
            If scores.Length >= 2 Then
                oFixture.MatchScore = Trim(scores(0).innerText) & "~" & Trim(scores(1).innerText)
                oFixture.MatchCardUrl = "https://tabletennis365.com" + Mid(doc.getElementsByClassName("matchCardIcon")(0).Children(0).href, 7)
            End If
            
            oFixtures.Add oFixture
        End If
    Next

    Set ParseWeekFixtures = oFixtures
End Function

Public Function ParseWeekNumber(title As String) As Integer
    Dim startPos, endPos As Integer
    
    startPos = InStr(title, " No ")
    If (startPos <= 0) Then
        ParseWeekNumber = 0
    Else
        endPos = InStr(startPos + 4, title, "-")
        If (endPos <= 0) Then
            ParseWeekNumber = CInt(Mid(title, startPos + 4))
        Else
            ParseWeekNumber = CInt(Mid(title, startPos + 4, endPos - startPos - 4))
        End If
    End If

End Function
