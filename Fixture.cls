VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: Fixture

Private m_iWeekNumber As Integer
Private m_dMatchDate As Date
Private m_sHomeTeam As String
Private m_sAwayTeam As String
Private m_sMatchScore As String
Private m_sMatchCardUrl As String

Public Property Get WeekNumber() As Integer
    WeekNumber = m_iWeekNumber
End Property

Public Property Let WeekNumber(ByVal iWeekNumber As Integer)
    m_iWeekNumber = iWeekNumber
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

Public Property Get MatchCardUrl() As String
    MatchCardUrl = m_sMatchCardUrl
End Property

Public Property Let MatchCardUrl(ByVal sMatchCardUrl As String)
    m_sMatchCardUrl = sMatchCardUrl
End Property

Public Function IsEquivalent(value As Fixture) As Boolean
    Dim result As Boolean
    result = WeekNumber = value.WeekNumber
    result = result And MatchDate = value.MatchDate
    result = result And HomeTeam = value.HomeTeam
    result = result And AwayTeam = value.AwayTeam
    result = result And MatchScore = value.MatchScore
    result = result And MatchCardUrl = value.MatchCardUrl
    IsEquivalent = result
End Function
