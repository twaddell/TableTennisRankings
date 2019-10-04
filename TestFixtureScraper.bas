Attribute VB_Name = "TestFixtureScraper"
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

Private Function GetTestFixtures() As HTMLDocument
    Dim html As String
    
    html = "<div id=""Fixtures""class=""fixtures divStyle""><div class=""caption"">Fixtures&gt;Wight-7pm</div><div class=""fixtureWeek""><div class=""title""><span>Match No 1-01 May 2019</span></div><div class=""week"">" _
+ "<div class=""fixture""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Blackgangers v's No Match at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-05-01"">Wednesday 1st May 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Blackgangers""itemprop=""isBasedOnUrl""title=""View team statistics for 'Blackgangers'"">Blackgangers</a></span></div></div></div><div class=""spacer""><div>vs</div></div><meta itemprop=""eventStatus""content=""http://schema.org/EventScheduled"">" _
+ "<div class=""away""><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/No_Match""itemprop=""isBasedOnUrl""title=""View team statistics for 'No Match'"">No Match</a></span></div></div></div><div class=""space-line0""></div></div>"
    
    html = html + "<div class=""fixture complete""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Dinosaurs v's Wobbly Surfers at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-05-01"">Wednesday 1st May 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344420""itemprop=""isBasedOnUrl""title=""View Match Card"">4</a></div><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Dinosaurs""itemprop=""isBasedOnUrl""title=""View team statistics for 'Dinosaurs'"">Dinosaurs</a></span></div>" _
+ "<div class=""playerName pom""title=""Player of the match""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Ricky_Lock/291064""title=""View player statistics"">Ricky Lock</a></span>(2)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Conal_Howells/291065""title=""View player statistics"">Conal Howells</a></span>(2)</div></div></div><div class=""spacer""><div>vs</div><div class=""matchCardIcon""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344420""><img class=""matchCard""src=""/Content/Themes/London/Images/icons/matchCard.png""title=""View matchcard between: Dinosaurs and Wobbly Surfers""></a></div></div>" _
+ "<div class=""away""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344420""itemprop=""isBasedOnUrl""title=""View Match Card"">1</a></div><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Wobbly_Surfers""itemprop=""isBasedOnUrl""title=""View team statistics for 'Wobbly Surfers'"">Wobbly Surfers</a></span></div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Liz_Arnold/291086""title=""View player statistics"">Liz Arnold</a></span>(0)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Rod_Medway/291087""title=""View player statistics"">Rod Medway</a></span>(0)</div></div></div><div class=""space-line0""></div></div>"

html = html + "<div class=""fixture complete""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Chessel Potters v's Jet Skiers at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-05-01"">Wednesday 1st May 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344430""itemprop=""isBasedOnUrl""title=""View Match Card"">2</a></div><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Chessel_Potters""itemprop=""isBasedOnUrl""title=""View team statistics for 'Chessel Potters'"">Chessel Potters</a></span></div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Roger_Nevell/291492""title=""View player statistics"">Roger Nevell</a></span>(1)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Pauline_Rorke/291067""title=""View player statistics"">Pauline Rorke</a></span>(1)</div></div></div><div class=""spacer""><div>vs</div><div class=""matchCardIcon""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344430""><img class=""matchCard""src=""/Content/Themes/London/Images/icons/matchCard.png""title=""View matchcard between: Chessel Potters and Jet Skiers""></a></div></div>" _
+ "<div class=""away""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344430""itemprop=""isBasedOnUrl""title=""View Match Card"">3</a></div><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Jet_Skiers""itemprop=""isBasedOnUrl""title=""View team statistics for 'Jet Skiers'"">Jet Skiers</a></span></div>" _
+ "<div class=""playerName pom""title=""Player of the match""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Peter_Dove/291493""title=""View player statistics"">Peter Dove</a></span>(2)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Martin_Dove/291056""title=""View player statistics"">Martin Dove</a></span>(0)</div></div></div><div class=""space-line0""></div></div>"

html = html + "<div class=""fixture complete""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Bovver Bats v's Old Fossils at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-05-01"">Wednesday 1st May 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344421""itemprop=""isBasedOnUrl""title=""View Match Card"">2</a></div><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Bovver_Bats""itemprop=""isBasedOnUrl""title=""View team statistics for 'Bovver Bats'"">Bovver Bats</a></span></div>" _
+ "<div class=""playerName pom""title=""Player of the match"" itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Alan_Hulbert/291068""title=""View player statistics"">Alan Hulbert</a></span>(2)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Karen_King/291069""title=""View player statistics"">Karen King</a></span>(0)</div></div></div><div class=""spacer""><div>vs</div><div class=""matchCardIcon""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344421""><img class=""matchCard""src=""/Content/Themes/London/Images/icons/matchCard.png""title=""View matchcard between: Bovver Bats and Old Fossils""></a></div></div>" _
+ "<div class=""away""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344421""itemprop=""isBasedOnUrl""title=""View Match Card"">3</a></div><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Old_Fossils""itemprop=""isBasedOnUrl""title=""View team statistics for 'Old Fossils'"">Old Fossils</a></span></div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Alan_Burch/291082""title=""View player statistics"">Alan Burch</a></span>(1)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Peter_Green/291084""title=""View player statistics"">Peter Green</a></span>(1)</div></div></div><div class=""space-line0""></div></div>"

html = html + "<div class=""fixture complete""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Go-karters v's Roman Villans at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-05-01"">Wednesday 1st May 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344424""itemprop=""isBasedOnUrl""title=""View Match Card"">3</a></div><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Go-karters""itemprop=""isBasedOnUrl""title=""View team statistics for 'Go-karters'"">Go-karters</a></span></div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/David_Batten/291074""title=""View player statistics"">David Batten</a></span>(1)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Ross_Holme/291075""title=""View player statistics"">Ross Holme</a></span>(1)</div></div></div><div class=""spacer""><div>vs</div><div class=""matchCardIcon""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344424""><img class=""matchCard""src=""/Content/Themes/London/Images/icons/matchCard.png""title=""View matchcard between: Go-karters and Roman Villans""></a></div></div>" _
+ "<div class=""away""><div class=""score""><a href=""/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344424""itemprop=""isBasedOnUrl""title=""View Match Card"">2</a></div><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Roman_Villans""itemprop=""isBasedOnUrl""title=""View team statistics for 'Roman Villans'"">Roman Villans</a></span></div>" _
+ "<div class=""playerName pom""title=""Player of the match""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Matt_Mair/291076""title=""View player statistics"">Matt Mair</a></span>(1)</div>" _
+ "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Summer_2019/Will_Henderson/291077""title=""View player statistics"">Will Henderson</a></span>(1)</div></div></div><div class=""space-line0""></div></div></div></div>" _

html = html + "<div class=""fixtureWeek""><div class=""title""><span>Match No 13-24 Jul 2019</span></div><div class=""week"">" _

html = html + "<div class=""fixture""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Blackgangers v's Dinosaurs at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-07-24"">Wednesday 24th July 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Blackgangers""itemprop=""isBasedOnUrl""title=""View team statistics for 'Blackgangers'"">Blackgangers</a></span></div></div></div><div class=""spacer""><div>vs</div></div><meta itemprop=""eventStatus""content=""http://schema.org/EventScheduled"">""" _
+ "<div class=""away""><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Dinosaurs""itemprop=""isBasedOnUrl""title=""View team statistics for 'Dinosaurs'"">Dinosaurs</a></span></div></div></div><div class=""space-line0""></div></div>"

html = html + "<div class=""fixture""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Deckchair Dollies v's Go-karters at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-07-24"">Wednesday 24th July 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Deckchair_Dollies""itemprop=""isBasedOnUrl""title=""View team statistics for 'Deckchair Dollies'"">Deckchair Dollies</a></span></div></div></div><div class=""spacer""><div>vs</div></div><meta itemprop=""eventStatus""content=""http://schema.org/EventScheduled"">" _
+ "<div class=""away""><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Go-karters""itemprop=""isBasedOnUrl""title=""View team statistics for 'Go-karters'"">Go-karters</a></span></div></div></div><div class=""space-line0""></div></div>"

html = html + "<div class=""fixture""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Shipwrecks v's Railway Steamers at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-07-24"">Wednesday 24th July 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Shipwrecks""itemprop=""isBasedOnUrl""title=""View team statistics for 'Shipwrecks'"">Shipwrecks</a></span></div></div></div><div class=""spacer""><div>vs</div></div><meta itemprop=""eventStatus""content=""http://schema.org/EventScheduled"">" _
+ "<div class=""away""><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Railway_Steamers""itemprop=""isBasedOnUrl""title=""View team statistics for 'Railway Steamers'"">Railway Steamers</a></span></div></div></div><div class=""space-line0""></div></div>"

html = html + "<div class=""fixture""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Old Fossils v's Wight Wiff-waffs at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-07-24"">Wednesday 24th July 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Old_Fossils""itemprop=""isBasedOnUrl""title=""View team statistics for 'Old Fossils'"">Old Fossils</a></span></div></div></div><div class=""spacer""><div>vs</div></div><meta itemprop=""eventStatus""content=""http://schema.org/EventScheduled"">" _
+ "<div class=""away""><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Wight_Wiff-waffs""itemprop=""isBasedOnUrl""title=""View team statistics for 'Wight Wiff-waffs'"">Wight Wiff-waffs</a></span></div></div></div><div class=""space-line0""></div></div>"

html = html + "<div class=""fixture""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Jet Skiers v's Bovver Bats at yj0kluubef""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2019-07-24"">Wednesday 24th July 2019</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
+ "<div class=""home""><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Jet_Skiers""itemprop=""isBasedOnUrl""title=""View team statistics for 'Jet Skiers'"">Jet Skiers</a></span></div></div></div><div class=""spacer""><div>vs</div></div><meta itemprop=""eventStatus""content=""http://schema.org/EventScheduled"">" _
+ "<div class=""away""><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Summer_2019/Wight_-_7pm/Bovver_Bats""itemprop=""isBasedOnUrl""title=""View team statistics for 'Bovver Bats'"">Bovver Bats</a></span></div></div></div><div class=""space-line0""></div></div>"
    
    Dim doc As HTMLDocument
    Set doc = New HTMLDocument
    doc.body.innerHTML = html
    Set GetTestFixtures = doc
End Function

'@TestMethod("Uncategorized")
Private Sub ParseFixturesShouldReturnFixturesCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim oFixtureHelper As FixtureHelper
    Set oFixtureHelper = New FixtureHelper
    
    Dim expected, actual As Collection
    Dim service As FixtureScraper
    Dim doc As HTMLDocument
    
    Set expected = New Collection
    
    expected.Add oFixtureHelper.CreateFixture(1, DateValue("2019-05-01"), "Blackgangers", "No Match", "", "")
    expected.Add oFixtureHelper.CreateFixture(1, DateValue("2019-05-01"), "Dinosaurs", "Wobbly Surfers", "4~1", "https://tabletennis365.com/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344420")
    expected.Add oFixtureHelper.CreateFixture(1, DateValue("2019-05-01"), "Chessel Potters", "Jet Skiers", "2~3", "https://tabletennis365.com/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344430")
    expected.Add oFixtureHelper.CreateFixture(1, DateValue("2019-05-01"), "Bovver Bats", "Old Fossils", "2~3", "https://tabletennis365.com/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344421")
    expected.Add oFixtureHelper.CreateFixture(1, DateValue("2019-05-01"), "Go-karters", "Roman Villans", "3~2", "https://tabletennis365.com/IsleOfWight/Results/Summer_2019/Wight_-_7pm/MatchCard/344424")
    expected.Add oFixtureHelper.CreateFixture(13, DateValue("2019-07-24"), "Blackgangers", "Dinosaurs", "", "")
    expected.Add oFixtureHelper.CreateFixture(13, DateValue("2019-07-24"), "Deckchair Dollies", "Go-karters", "", "")
    expected.Add oFixtureHelper.CreateFixture(13, DateValue("2019-07-24"), "Shipwrecks", "Railway Steamers", "", "")
    expected.Add oFixtureHelper.CreateFixture(13, DateValue("2019-07-24"), "Old Fossils", "Wight Wiff-waffs", "", "")
    expected.Add oFixtureHelper.CreateFixture(13, DateValue("2019-07-24"), "Jet Skiers", "Bovver Bats", "", "")
            
    Set doc = GetTestFixtures()
    
    'Act:
    Set service = New FixtureScraper
    Set actual = service.ParseFixtures(doc)
    
    'Assert:
    Assert.AreEqual expected.Count, actual.Count
    
    Dim i As Integer
    For i = 1 To expected.Count
        Assert.IsTrue expected(i).IsEquivalent(actual(i)), "actual " + CStr(i) + " did not equal expected"
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseFixturesShouldSkipFreeWeeks()
    On Error GoTo TestFail
    
    'Arrange:
    Dim oFixtureHelper As FixtureHelper
    Set oFixtureHelper = New FixtureHelper
    
    Dim expected, actual As Collection
    Dim service As FixtureScraper
    
    Set expected = New Collection
    
    expected.Add oFixtureHelper.CreateFixture(5, DateValue("2018-11-01"), "Ryde Run Of The Mill", "Ryde A", "1~9", "https://tabletennis365.com/IsleOfWight/Results/Winter_2018-19/Division_1/MatchCard/301874")
            
    Dim html As String
    
    html = "<div id=""Fixtures""class=""fixtures divStyle""><div class=""caption"">Fixtures&gt;Division 1</div>" _
    + "<div class=""fixtureWeek""><div class=""title""><span>Week No 4</span></div><div class=""week""><div class=""fixtureSkip"">**  Free </div></div></div>"

    html = html + "<div class=""fixtureWeek""><div class=""title""><span>Week No 5-30 Oct 2018&gt;01 Nov 2018</span></div><div class=""week""><div class=""fixture complete""itemscope=""""itemtype=""http://schema.org/SportsEvent""><meta itemprop=""description""content=""Ryde Run Of The Mill v's Ryde A at Ryde Run Of The Mill""><meta itemprop=""organizer""itemscope=""""itemtype=""http://schema.org/Organization""itemref=""Org-Isle of Wight Table Tennis League""><div class=""date""itemprop=""startDate""><time datetime=""2018-11-01"">Thursday 1st November 2018</time></div><div class=""venue""itemprop=""location""><span itemscope=""""itemtype=""http://schema.org/SportsActivityLocation""itemprop=""name""><a href="""">Ryde TTC</a></span></div><div class=""venueOptionalInfo""></div>" _
    + "<div class=""home""><div class=""score""><a href=""/IsleOfWight/Results/Winter_2018-19/Division_1/MatchCard/301874""itemprop=""isBasedOnUrl""title=""View Match Card"">1</a></div><div class=""homeTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Winter_2018-19/Division_1/Ryde_Run_Of_The_Mill""itemprop=""isBasedOnUrl""title=""View team statistics for 'Ryde Run Of The Mill'"">Ryde Run Of The Mill</a></span></div>" _
    + "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/John_Cornforth/269496""title=""View player statistics"">John Cornforth</a></span>(0)</div>" _
    + "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Mark_Holbrook/269502""title=""View player statistics"">Mark Holbrook</a></span>(0)</div>" _
    + "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Ed_Kennedy/269498""title=""View player statistics"">Ed Kennedy</a></span>(1)</div></div></div><div class=""spacer""><div>vs</div><div class=""matchCardIcon""><a href=""/IsleOfWight/Results/Winter_2018-19/Division_1/MatchCard/301874""><img class=""matchCard""src=""/Content/Themes/London/Images/icons/matchCard.png""title=""View matchcard between: Ryde Run Of The Mill and Ryde A""></a></div></div>" _
    + "<div class=""away""><div class=""score""><a href=""/IsleOfWight/Results/Winter_2018-19/Division_1/MatchCard/301874""itemprop=""isBasedOnUrl""title=""View Match Card"">9</a></div><div class=""awayTeam""><div class=""teamName ""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/SportsTeam""itemprop=""name""><a href=""/IsleOfWight/Results/Team/Statistics/Winter_2018-19/Division_1/Ryde_A""itemprop=""isBasedOnUrl""title=""View team statistics for 'Ryde A'"">Ryde A</a></span></div>" _
    + "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Lee_Orton/269477""title=""View player statistics"">Lee Orton</a></span>(3)</div>" _
    + "<div class=""playerName pom""title=""Player of the match""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Ollie_Staniforth/269479""title=""View player statistics"">Ollie Staniforth</a></span>(3)</div>" _
    + "<div class=""playerName""title=""""itemprop=""performer""><span itemscope=""""itemtype=""http://schema.org/Person""itemprop=""name""><a href=""/IsleOfWight/Results/Player/Statistics/Winter_2018-19/Roger_Nevell/269475""title=""View player statistics"">Roger Nevell</a></span>(2)</div></div></div><div class=""space-line0""></div></div></div></div>"
        
    Dim doc As HTMLDocument
    Set doc = New HTMLDocument
    doc.body.innerHTML = html
    
    'Act:
    Set service = New FixtureScraper
    Set actual = service.ParseFixtures(doc)
    
    'Assert:
    Assert.AreEqual expected.Count, actual.Count
    
    Dim i As Integer
    For i = 1 To expected.Count
        Assert.IsTrue expected(i).IsEquivalent(actual(i)), "actual " + CStr(i) + " did not equal expected"
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseWeekNumberShouldGetCorrectWeekNumberForFreeWeeks()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected, actual As Integer
    Dim service As FixtureScraper
    
    expected = 4
    
    'Act:
    Set service = New FixtureScraper
    actual = service.ParseWeekNumber("Week No 4")
    
    'Assert:
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseWeekNumberShouldGetCorrectWeekNumberForMultiDateWeek()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected, actual As Integer
    Dim service As FixtureScraper
    
    expected = 5
    
    'Act:
    Set service = New FixtureScraper
    actual = service.ParseWeekNumber("Week No 5 - 30 Oct 2018 > 01 Nov 2018")
    
    'Assert:
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub ParseWeekNumberShouldGetCorrectWeekNumberForSingleDateWeek()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected, actual As Integer
    Dim service As FixtureScraper
    
    expected = 3
    
    'Act:
    Set service = New FixtureScraper
    actual = service.ParseWeekNumber("Match No 3 - 15 May 2019")
    
    'Assert:
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
