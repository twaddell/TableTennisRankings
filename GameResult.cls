VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: GameResult

Private m_sWinner As String
Private m_oEnds() As String

Private Sub Class_Initialize()
    ReDim m_oEnds(4)
End Sub

Public Property Get Ends(index As Long) As String
    Ends = m_oEnds(index)
End Property

Public Property Let Ends(index As Long, value As String)
    If index > UBound(m_oEnds) Then ReDim Preserve m_oEnds(index)
    m_oEnds(index) = value
End Property

Public Property Get Winner() As String
    Winner = m_sWinner
End Property

Public Property Let Winner(ByVal sWinner As String)
    m_sWinner = sWinner
End Property

Public Function IsEquivalent(value As GameResult) As Boolean
    Dim result As Boolean
    result = Winner = value.Winner
    result = result And Ends(0) = value.Ends(0)
    result = result And Ends(1) = value.Ends(1)
    result = result And Ends(2) = value.Ends(2)
    result = result And Ends(3) = value.Ends(3)
    result = result And Ends(4) = value.Ends(4)
    IsEquivalent = result
End Function
