VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TournamentResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: TournamentResult

Private m_dMatchDate As Date
Private m_sPlayerA As String
Private m_sPlayerB As String
Private m_sWinner As String

' Properties
Public Property Get MatchDate() As Date
    MatchDate = m_dMatchDate
End Property

Public Property Let MatchDate(ByVal dMatchDate As Date)
    m_dMatchDate = dMatchDate
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

Public Property Get Winner() As String
    Winner = m_sWinner
End Property

Public Property Let Winner(ByVal sWinner As String)
    m_sWinner = sWinner
End Property
