VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_New()
    Call PCC_Footer
End Sub

Private Sub Document_Open()
    If Not ActiveDocument.Type = wdTypeTemplate Then
        Call PCC_Footer
    End If
End Sub
