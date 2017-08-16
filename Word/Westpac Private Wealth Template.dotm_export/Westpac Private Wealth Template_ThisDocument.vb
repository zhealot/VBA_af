VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    If ContentControl.Title = "ccRiskProfile" Then
        DocPrty "RiskProfile", ContentControl.Range.Text
        ActiveDocument.Content.Fields.Update
    End If
End Sub

Private Sub Document_New()
    frmMain.Show
End Sub
