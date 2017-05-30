VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cb As MSForms.OptionButton
Attribute cb.VB_VarHelpID = -1

Public Sub cb_Click()
    'MsgBox "hdshds"
End Sub

Public Sub cb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'MsgBox "Dbl"
End Sub
