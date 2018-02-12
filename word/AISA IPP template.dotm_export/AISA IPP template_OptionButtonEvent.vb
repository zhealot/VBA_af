VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionButtonEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents oOBEvents As MSForms.OptionButton
Attribute oOBEvents.VB_VarHelpID = -1

Private Sub oOBEvents_Click()
    If oOBEvents.Caption = "None of these things" Then
        frmMain.CommandButton1.Caption = "Next"
    Else
        frmMain.CommandButton1.Caption = "Yes"
    End If
End Sub
