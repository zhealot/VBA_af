VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents oButton As MSForms.CommandButton
Attribute oButton.VB_VarHelpID = -1

Private Sub oButton_Click()
    MsgBox "button clicked"
End Sub
