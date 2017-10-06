VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmWelcome 
   Caption         =   "ASIA IPP Template"
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   OleObjectBlob   =   "AISA IPP template_fmWelcome.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNext_Click()
    Me.Hide
    frmMain.Show
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    ThisDocument.FollowHyperlink "http://www.allfields.com"
End Sub
