VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Public WithEvents oApp As Word.Application
Attribute oApp.VB_VarHelpID = -1

Private Sub oApp_DocumentOpen(ByVal doc As Document)
    If Not ActiveDocument.Type = wdTypeTemplate Then
        Call PCC_Footer
    End If
End Sub

Private Sub oApp_NewDocument(ByVal doc As Document)
    Call PCC_Footer
End Sub
