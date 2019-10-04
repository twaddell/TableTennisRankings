VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressDialog 
   Caption         =   "Running..."
   ClientHeight    =   1035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ProgressDialog.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    SetProgress "Processing...", 0
End Sub

Public Sub SetProgress(sMessage As String, pct As Single)
    Message.Caption = sMessage
    ProgressBar.Width = pct * (ProgressFrame.Width)
    DoEvents
End Sub
