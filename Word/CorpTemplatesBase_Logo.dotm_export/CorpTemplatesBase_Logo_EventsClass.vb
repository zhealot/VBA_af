VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents oWdApp As Word.Application
Attribute oWdApp.VB_VarHelpID = -1

Private Sub oWdApp_DocumentOpen(ByVal Doc As Document)
'    MsgBox "opeeeeen"
    Call SetLogo(Doc)
End Sub
