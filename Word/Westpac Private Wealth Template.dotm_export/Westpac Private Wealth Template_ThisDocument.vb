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
        Dim sRisk As String
        Dim rg As Range
        Dim rgTmp As Range
        
        'autotext for Executive Summary part
        If ActiveDocument.Bookmarks.Exists("RiskProfile1") Then
            Set rg = ActiveDocument.Bookmarks("RiskProfile1").Range
            Err.Clear
            On Error Resume Next
            sRisk = "RP" & Replace(ContentControl.Range.Text, " ", "")
            Set rgTmp = ActiveDocument.AttachedTemplate.BuildingBlockEntries(sRisk).Insert(rg, True)
            If Err.Number <> 0 Then
                rg.Text = "Insert Risk Profile auto text"
                Set rgTmp = rg
            End If
            ActiveDocument.Bookmarks.Add "RiskProfile1", rgTmp
        End If
        'autotext for Risk Profile - Risk vs Return trade off part
        If ActiveDocument.Bookmarks.Exists("RiskProfile2") Then
            Set rg = ActiveDocument.Bookmarks("RiskProfile2").Range
            Err.Clear
            On Error Resume Next
            sRisk = "RiskProfile" & Left(ContentControl.Range.Text, InStr(ContentControl.Range.Text, " ") - 1)
            Set rgTmp = ActiveDocument.AttachedTemplate.BuildingBlockEntries(sRisk).Insert(rg, True)
            If Err.Number <> 0 Then
                rg.Text = "Insert Risk Profile auto text"
                Set rgTmp = rg
            End If
            ActiveDocument.Bookmarks.Add "RiskProfile2", rgTmp
        End If
        ActiveDocument.Content.Fields.Update
    End If
End Sub

Private Sub Document_New()
    frmMain.Show
End Sub
