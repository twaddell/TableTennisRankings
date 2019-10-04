Attribute VB_Name = "TestMatchScraper"
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

Private Function GetTestMatchCard() As HTMLDocument

    Dim doc As HTMLDocument
    Set doc = New HTMLDocument
    doc.body.innerHTML = "<div itemprop=""articleBody""id=""MatchCardBody""><div id=""PublicMatchCardTypeB""><div id=""CardSummary""class=""divStyle""><div class=""caption""><a href=""/IsleOfWight/Tables/Winter%202018-19/Division%202"">Division 2</a>&gt;Match Card</div><div class=""container""><div class=""topBar cardSummary""><div class=""menuItem menuLeft""><div class=""title"">Fixture Details</div><div class=""teamNames""><span><a href=""/IsleOfWight/Results/Team/Statistics/Winter_2018-19/Division_2/Shanklin_B""itemprop=""isBasedOnUrl""title=""View team statistics for 'Shanklin B'"">Shanklin B</a></span>v<span><a href=""/IsleOfWight/Results/Team/Statistics/Winter_2018-19/Division_2/Ryde_Rascals""itemprop=""isBasedOnUrl""title=""View team statistics for 'Ryde Rascals'"">Ryde Rascals</a></span></div><div class=""matchNo"">Match Number:13</div><div class=""dates"">Match Date:<time datetime=""2019-01-31"">31 Jan 2019</time></div><div class=""matchScore"">Match Score:8-2</div></div>" _
+ "<div class=""menuItem menuRight playerOfTheMatch""><div class=""title"">Player Of The Match</div><div class=""player""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Darol_Wilson/269557""title=""View player statistics"">Darol Wilson</a></div></div><div class=""space-line0""></div></div><div class=""space-line0""></div><div id=""CardResults""><div class=""table-row rowX row1""><div class=""col1"">&nbsp;</div><div class=""col2""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Alison_Botha/269534""title=""View player statistics"">Alison Botha</a></div><div class=""col3""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Peter_Clarke/269533""title=""View player statistics"">Peter Clarke</a></div><div class=""col4""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/John_Featherstone/269535""title=""View player statistics"">John Featherstone</a></div></div><div class=""space-line0""></div>" _
+ "<div class=""table-row rowX row2""><div class=""col1""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Alan_Dennis/269556""title=""View player statistics"">Alan Dennis</a></div><div class=""set""><div class=""game1""><div class=""gameScore"">11-7</div></div><div class=""game2""><div class=""gameScore"">11-7</div></div><div class=""game3""><div class=""gameScore"">8-11</div></div><div class=""game4""><div class=""gameScore"">11-8</div></div></div><div class=""set""><div class=""game1""><div class=""gameScore"">11-7</div></div><div class=""game2""><div class=""gameScore"">3-11</div></div><div class=""game3""><div class=""gameScore"">12-14</div></div><div class=""game4""><div class=""gameScore"">8-11</div></div></div><div class=""set""><div class=""game1""><div class=""gameScore"">11-8</div></div><div class=""game2""><div class=""gameScore"">10-12</div></div><div class=""game3""><div class=""gameScore"">6-11</div></div><div class=""game4""><div class=""gameScore"">11-9</div></div>" _
+ "<div class=""game5""><div class=""gameScore"">11-4</div></div></div></div><div class=""space-line0""></div><div class=""table-row rowX row3""><div class=""col1""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Darol_Wilson/269557""title=""View player statistics"">Darol Wilson</a></div><div class=""set""><div class=""game1""><div class=""gameScore"">11-6</div></div><div class=""game2""><div class=""gameScore"">6-11</div></div><div class=""game3""><div class=""gameScore"">11-4</div></div><div class=""game4""><div class=""gameScore"">3-11</div></div><div class=""game5""><div class=""gameScore"">11-7</div></div></div><div class=""set""><div class=""game1""><div class=""gameScore"">11-5</div></div><div class=""game2""><div class=""gameScore"">11-1</div></div><div class=""game3""><div class=""gameScore"">11-8</div></div></div>" _
+ "<div class=""set""><div class=""game1""><div class=""gameScore"">5-11</div></div><div class=""game2""><div class=""gameScore"">13-11</div></div><div class=""game3""><div class=""gameScore"">11-6</div></div><div class=""game4""><div class=""gameScore"">14-16</div></div><div class=""game5""><div class=""gameScore"">11-8</div></div></div></div><div class=""space-line0""></div><div class=""table-row rowX row4""><div class=""col1""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Tom_Waddell/269555""title=""View player statistics"">Tom Waddell</a></div><div class=""set""><div class=""game1""><div class=""gameScore"">8-11</div></div><div class=""game2""><div class=""gameScore"">11-6</div></div><div class=""game3""><div class=""gameScore"">8-11</div></div><div class=""game4""><div class=""gameScore"">11-8</div></div><div class=""game5""><div class=""gameScore"">11-9</div></div></div>" _
+ "<div class=""set""><div class=""game1""><div class=""gameScore"">11-5</div></div><div class=""game2""><div class=""gameScore"">11-9</div></div><div class=""game3""><div class=""gameScore"">10-12</div></div><div class=""game4""><div class=""gameScore"">9-11</div></div><div class=""game5""><div class=""gameScore"">11-9</div></div></div><div class=""set""><div class=""game1""><div class=""gameScore"">8-11</div></div><div class=""game2""><div class=""gameScore"">6-11</div></div><div class=""game3""><div class=""gameScore"">8-11</div></div></div></div><div class=""space-line0""></div><div class=""table-row rowX row5""><div class=""col1""><div><div class=""dPlayer""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Alan_Dennis/269556""title=""View player statistics"">Alan Dennis</a><br><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Darol_Wilson/269557""title=""View player statistics"">Darol Wilson</a></div></div>" _
+ "<div><div><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Alison_Botha/269534""title=""View player statistics"">Alison Botha</a><br><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/John_Featherstone/269535""title=""View player statistics"">John Featherstone</a></div></div></div><div class=""doublesSet""><div class=""game1""><div class=""gameScore"">11-8</div></div><div class=""game2""><div class=""gameScore"">11-9</div></div><div class=""game3""><div class=""gameScore"">7-11</div></div><div class=""game4""><div class=""gameScore"">11-6</div></div></div><div class=""authorised""><div>Submitted By:Edward Purvey</div><div>Approved By:Edward Purvey</div><div>Completed By:Edward Purvey</div></div></div><div class=""space-line0""></div></div></div></div></div></div>"

    Set GetTestMatchCard = doc
End Function

'@TestMethod("Uncategorized")
Private Sub ParseMatchCardShouldReturnMatchResult()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As MatchResult
    Dim actual As MatchResult
    Dim service As MatchScraper
    Dim doc As HTMLDocument
    
    Set expected = New MatchResult
    expected.Key = ""
    expected.WeekNumber = 0
    expected.Format = "9S1D"
    
    expected.MatchDate = DateValue("31-Jan-19")
    expected.HomeTeam = "Shanklin B"
    expected.AwayTeam = "Ryde Rascals"
    expected.MatchScore = "8~2"
    expected.PlayerA = "Alan Dennis"
    expected.PlayerB = "Darol Wilson"
    expected.PlayerC = "Tom Waddell"
    expected.PlayerX = "Alison Botha"
    expected.PlayerY = "Peter Clarke"
    expected.PlayerZ = "John Featherstone"
    expected.Games(0).Ends(0) = "11~7"
    expected.Games(0).Ends(1) = "11~7"
    expected.Games(0).Ends(2) = "8~11"
    expected.Games(0).Ends(3) = "11~8"
    expected.Games(0).Ends(4) = ""
    expected.Games(0).Winner = "Home"
    expected.Games(1).Ends(0) = "11~5"
    expected.Games(1).Ends(1) = "11~1"
    expected.Games(1).Ends(2) = "11~8"
    expected.Games(1).Ends(3) = ""
    expected.Games(1).Ends(4) = ""
    expected.Games(1).Winner = "Home"
    expected.Games(2).Ends(0) = "8~11"
    expected.Games(2).Ends(1) = "6~11"
    expected.Games(2).Ends(2) = "8~11"
    expected.Games(2).Ends(3) = ""
    expected.Games(2).Ends(4) = ""
    expected.Games(2).Winner = "Away"
    expected.Games(3).Ends(0) = "11~6"
    expected.Games(3).Ends(1) = "6~11"
    expected.Games(3).Ends(2) = "11~4"
    expected.Games(3).Ends(3) = "3~11"
    expected.Games(3).Ends(4) = "11~7"
    expected.Games(3).Winner = "Home"
    expected.Games(4).Ends(0) = "11~8"
    expected.Games(4).Ends(1) = "10~12"
    expected.Games(4).Ends(2) = "6~11"
    expected.Games(4).Ends(3) = "11~9"
    expected.Games(4).Ends(4) = "11~4"
    expected.Games(4).Winner = "Home"
    expected.Games(5).Ends(0) = "11~5"
    expected.Games(5).Ends(1) = "11~9"
    expected.Games(5).Ends(2) = "10~12"
    expected.Games(5).Ends(3) = "9~11"
    expected.Games(5).Ends(4) = "11~9"
    expected.Games(5).Winner = "Home"
    expected.Games(6).Ends(0) = "5~11"
    expected.Games(6).Ends(1) = "13~11"
    expected.Games(6).Ends(2) = "11~6"
    expected.Games(6).Ends(3) = "14~16"
    expected.Games(6).Ends(4) = "11~8"
    expected.Games(6).Winner = "Home"
    expected.Games(7).Ends(0) = "8~11"
    expected.Games(7).Ends(1) = "11~6"
    expected.Games(7).Ends(2) = "8~11"
    expected.Games(7).Ends(3) = "11~8"
    expected.Games(7).Ends(4) = "11~9"
    expected.Games(7).Winner = "Home"
    expected.Games(8).Ends(0) = "11~7"
    expected.Games(8).Ends(1) = "3~11"
    expected.Games(8).Ends(2) = "12~14"
    expected.Games(8).Ends(3) = "8~11"
    expected.Games(8).Ends(4) = ""
    expected.Games(8).Winner = "Away"
    expected.Games(9).Ends(0) = "11~8"
    expected.Games(9).Ends(1) = "11~9"
    expected.Games(9).Ends(2) = "7~11"
    expected.Games(9).Ends(3) = "11~6"
    expected.Games(9).Ends(4) = ""
    expected.Games(9).Winner = "Home"
        
    Set doc = GetTestMatchCard()
    
    'Act:
    Set service = New MatchScraper
    Set actual = service.ParseMatchCard(doc)
    
    'Assert:
    Assert.IsTrue expected.IsEquivalent(actual), "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParsePlayerNamesShouldPopulatePlayerNames()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As MatchResult
    Dim actual As MatchResult
    Dim service As MatchScraper
    Dim doc As HTMLDocument
    
    Set expected = New MatchResult
    expected.PlayerA = "Alan Dennis"
    expected.PlayerB = "Darol Wilson"
    expected.PlayerC = "Tom Waddell"
    expected.PlayerX = "Alison Botha"
    expected.PlayerY = "Peter Clarke"
    expected.PlayerZ = "John Featherstone"
            
    Set doc = GetTestMatchCard()
    
    'Act:
    Set service = New MatchScraper
    Set actual = New MatchResult
    Set actual = service.ParsePlayerNames(actual, doc)
    
    'Assert:
    Assert.IsTrue expected.IsEquivalent(actual), "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseMatchSummaryShouldPopulateMatchSummaryDetails()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As MatchResult
    Dim actual As MatchResult
    Dim service As MatchScraper
    Dim doc As HTMLDocument
    
    Set expected = New MatchResult
    expected.HomeTeam = "Shanklin B"
    expected.AwayTeam = "Ryde Rascals"
    expected.MatchDate = DateValue("31/01/2019")
    expected.MatchScore = "8~2"
            
    Set doc = GetTestMatchCard()
    
    'Act:
    Set service = New MatchScraper
    Set actual = New MatchResult
    Set actual = service.ParseMatchSummary(actual, doc)
    
    'Assert:
    Assert.IsTrue expected.IsEquivalent(actual), "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseSinglesGameScoresShouldPopulateGameScores()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As MatchResult
    Dim actual As MatchResult
    Dim service As MatchScraper
    Dim doc As HTMLDocument
    
    Set expected = New MatchResult
    
    expected.Games(0).Ends(0) = "11~7"
    expected.Games(0).Ends(1) = "11~7"
    expected.Games(0).Ends(2) = "8~11"
    expected.Games(0).Ends(3) = "11~8"
    expected.Games(0).Ends(4) = ""
    expected.Games(0).Winner = "Home"
    expected.Games(1).Ends(0) = "11~5"
    expected.Games(1).Ends(1) = "11~1"
    expected.Games(1).Ends(2) = "11~8"
    expected.Games(1).Ends(3) = ""
    expected.Games(1).Ends(4) = ""
    expected.Games(1).Winner = "Home"
    expected.Games(2).Ends(0) = "8~11"
    expected.Games(2).Ends(1) = "6~11"
    expected.Games(2).Ends(2) = "8~11"
    expected.Games(2).Ends(3) = ""
    expected.Games(2).Ends(4) = ""
    expected.Games(2).Winner = "Away"
    expected.Games(3).Ends(0) = "11~6"
    expected.Games(3).Ends(1) = "6~11"
    expected.Games(3).Ends(2) = "11~4"
    expected.Games(3).Ends(3) = "3~11"
    expected.Games(3).Ends(4) = "11~7"
    expected.Games(3).Winner = "Home"
    expected.Games(4).Ends(0) = "11~8"
    expected.Games(4).Ends(1) = "10~12"
    expected.Games(4).Ends(2) = "6~11"
    expected.Games(4).Ends(3) = "11~9"
    expected.Games(4).Ends(4) = "11~4"
    expected.Games(4).Winner = "Home"
    expected.Games(5).Ends(0) = "11~5"
    expected.Games(5).Ends(1) = "11~9"
    expected.Games(5).Ends(2) = "10~12"
    expected.Games(5).Ends(3) = "9~11"
    expected.Games(5).Ends(4) = "11~9"
    expected.Games(5).Winner = "Home"
    expected.Games(6).Ends(0) = "5~11"
    expected.Games(6).Ends(1) = "13~11"
    expected.Games(6).Ends(2) = "11~6"
    expected.Games(6).Ends(3) = "14~16"
    expected.Games(6).Ends(4) = "11~8"
    expected.Games(6).Winner = "Home"
    expected.Games(7).Ends(0) = "8~11"
    expected.Games(7).Ends(1) = "11~6"
    expected.Games(7).Ends(2) = "8~11"
    expected.Games(7).Ends(3) = "11~8"
    expected.Games(7).Ends(4) = "11~9"
    expected.Games(7).Winner = "Home"
    expected.Games(8).Ends(0) = "11~7"
    expected.Games(8).Ends(1) = "3~11"
    expected.Games(8).Ends(2) = "12~14"
    expected.Games(8).Ends(3) = "8~11"
    expected.Games(8).Ends(4) = ""
    expected.Games(8).Winner = "Away"
            
    Set doc = GetTestMatchCard()
    
    'Act:
    Set service = New MatchScraper
    Set actual = New MatchResult
    Set actual = service.ParseSinglesGameScores(actual, doc)
    
    'Assert:
    Assert.IsTrue expected.IsEquivalent(actual), "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseDoublesGameScoresShouldPopulateGameScores()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As MatchResult
    Dim actual As MatchResult
    Dim service As MatchScraper
    Dim doc As HTMLDocument
    
    Set expected = New MatchResult
    expected.Games(9).Ends(0) = "11~8"
    expected.Games(9).Ends(1) = "11~9"
    expected.Games(9).Ends(2) = "7~11"
    expected.Games(9).Ends(3) = "11~6"
    expected.Games(9).Ends(4) = ""
    expected.Games(9).Winner = "Home"
            
    Set doc = GetTestMatchCard()
    
    'Act:
    Set service = New MatchScraper
    Set actual = New MatchResult
    Set actual = service.ParseDoublesGameScores(actual, doc)
    
    'Assert:
    Assert.IsTrue expected.IsEquivalent(actual), "actual did not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

