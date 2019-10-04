VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: MatchResult

Private m_sKey As String
Private m_iWeekNumber As Integer
Private m_sFormat As String
Private m_sDivision As String
Private m_dMatchDate As Date
Private m_sHomeTeam As String
Private m_sAwayTeam As String
Private m_sMatchScore As String
Private m_sPlayerA As String
Private m_sPlayerB As String
Private m_sPlayerC As String
Private m_sPlayerX As String
Private m_sPlayerY As String
Private m_sPlayerZ As String
Private m_oGames() As GameResult
Private m_sComments As String
Private m_sTT365Date As String

Private Sub Class_Initialize()
    ReDim m_oGames(9)
    Set m_oGames(0) = New GameResult
    Set m_oGames(1) = New GameResult
    Set m_oGames(2) = New GameResult
    Set m_oGames(3) = New GameResult
    Set m_oGames(4) = New GameResult
    Set m_oGames(5) = New GameResult
    Set m_oGames(6) = New GameResult
    Set m_oGames(7) = New GameResult
    Set m_oGames(8) = New GameResult
    Set m_oGames(9) = New GameResult
End Sub

' Properties
Public Property Get Key() As String
    Key = m_sKey
End Property

Public Property Let Key(ByVal sKey As String)
    m_sKey = sKey
End Property

Public Property Get WeekNumber() As Integer
    WeekNumber = m_iWeekNumber
End Property

Public Property Let WeekNumber(ByVal sWeekNumber As Integer)
    m_iWeekNumber = sWeekNumber
End Property

Public Property Get Format() As String
    Format = m_sFormat
End Property

Public Property Let Format(ByVal sFormat As String)
    m_sFormat = sFormat
End Property

Public Property Get Division() As String
    Division = m_sDivision
End Property

Public Property Let Division(ByVal sDivision As String)
    m_sDivision = sDivision
End Property

Public Property Get MatchDate() As Date
    MatchDate = m_dMatchDate
End Property

Public Property Let MatchDate(ByVal dMatchDate As Date)
    m_dMatchDate = dMatchDate
End Property

Public Property Get HomeTeam() As String
    HomeTeam = m_sHomeTeam
End Property

Public Property Let HomeTeam(ByVal sHomeTeam As String)
    m_sHomeTeam = sHomeTeam
End Property

Public Property Get AwayTeam() As String
    AwayTeam = m_sAwayTeam
End Property

Public Property Let AwayTeam(ByVal sAwayTeam As String)
    m_sAwayTeam = sAwayTeam
End Property

Public Property Get MatchScore() As String
    MatchScore = m_sMatchScore
End Property

Public Property Let MatchScore(ByVal sMatchScore As String)
    m_sMatchScore = sMatchScore
End Property

Public Property Get PlayerA() As String
    PlayerA = m_sPlayerA
End Property

Public Property Let PlayerA(ByVal sPlayerA As String)
    m_sPlayerA = sPlayerA
End Property

Public Property Get PlayerB() As String
    PlayerB = m_sPlayerB
End Property

Public Property Let PlayerB(ByVal sPlayerB As String)
    m_sPlayerB = sPlayerB
End Property

Public Property Get PlayerC() As String
    PlayerC = m_sPlayerC
End Property

Public Property Let PlayerC(ByVal sPlayerC As String)
    m_sPlayerC = sPlayerC
End Property

Public Property Get PlayerX() As String
    PlayerX = m_sPlayerX
End Property

Public Property Let PlayerX(ByVal sPlayerX As String)
    m_sPlayerX = sPlayerX
End Property

Public Property Get PlayerY() As String
    PlayerY = m_sPlayerY
End Property

Public Property Let PlayerY(ByVal sPlayerY As String)
    m_sPlayerY = sPlayerY
End Property

Public Property Get PlayerZ() As String
    PlayerZ = m_sPlayerZ
End Property

Public Property Let PlayerZ(ByVal sPlayerZ As String)
    m_sPlayerZ = sPlayerZ
End Property

Public Property Get Games(ByVal index As Long) As GameResult
    Set Games = m_oGames(index)
End Property

Public Property Let Games(ByVal index As Long, ByVal value As GameResult)
    If index > UBound(m_oGames) Then ReDim Preserve m_oGames(index)
    Set m_oGames(index) = value
End Property

Public Property Get Comments() As String
    Comments = m_sComments
End Property

Public Property Let Comments(ByVal sComments As String)
    m_sComments = sComments
End Property

Public Property Get TT365Date() As String
    TT365Date = m_sTT365Date
End Property

Public Property Let TT365Date(ByVal sTT365Date As String)
    m_sTT365Date = sTT365Date
End Property

Public Function IsEquivalent(value As MatchResult) As Boolean
    Dim result As Boolean
    result = Key = value.Key
    result = result And Division = value.Division
    result = result And MatchDate = value.MatchDate
    result = result And WeekNumber = value.WeekNumber
    result = result And HomeTeam = value.HomeTeam
    result = result And AwayTeam = value.AwayTeam
    result = result And Format = value.Format
    result = result And Comments = value.Comments
    result = result And MatchScore = value.MatchScore
    result = result And TT365Date = value.TT365Date
    result = result And PlayerA = value.PlayerA
    result = result And PlayerB = value.PlayerB
    result = result And PlayerC = value.PlayerC
    result = result And PlayerX = value.PlayerX
    result = result And PlayerY = value.PlayerY
    result = result And PlayerZ = value.PlayerZ
    result = result And Games(0).IsEquivalent(value.Games(0))
    result = result And Games(1).IsEquivalent(value.Games(1))
    result = result And Games(2).IsEquivalent(value.Games(2))
    result = result And Games(3).IsEquivalent(value.Games(3))
    result = result And Games(4).IsEquivalent(value.Games(4))
    result = result And Games(5).IsEquivalent(value.Games(5))
    result = result And Games(6).IsEquivalent(value.Games(6))
    result = result And Games(7).IsEquivalent(value.Games(7))
    result = result And Games(8).IsEquivalent(value.Games(8))
    result = result And Games(9).IsEquivalent(value.Games(9))
    IsEquivalent = result
End Function

