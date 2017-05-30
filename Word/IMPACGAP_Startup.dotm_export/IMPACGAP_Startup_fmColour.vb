VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmColour 
   Caption         =   "Background colour"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2235
   OleObjectBlob   =   "IMPACGAP_Startup_fmColour.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "fmColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn1_Click()
    If Selection.Cells.Count > 0 Then
        Selection.Shading.BackgroundPatternColor = btn1.BackColor
        fmColour.Hide
    End If
End Sub

Private Sub btn2_Click()
    If Selection.Cells.Count > 0 Then
        Selection.Shading.BackgroundPatternColor = btn2.BackColor
        fmColour.Hide
    End If
End Sub

Private Sub btn3_Click()
    If Selection.Cells.Count > 0 Then
        Selection.Shading.BackgroundPatternColor = btn3.BackColor
        fmColour.Hide
    End If
End Sub

