VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "TEC Graphs"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
   OleObjectBlob   =   "Provider data with graphs v4 - Allfields_frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbProvider_Change()
    Call cbProvicderChange
End Sub

Private Sub cbSimilar_Change()
    If TriggerSimilar Then
        cbProvider.Value = cbSimilar.Value
    End If
End Sub

Private Sub CommandButton1_Click()
    Call Init(True)
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Call Init(True)
End Sub
