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
    Dim doc As Document
    Dim cc As ContentControl 'source ContentControl
    Dim ccTmp As ContentControl 'target ContentControl
        
    Set doc = ActiveDocument
    Set cc = ContentControl
    
    'disable paragraph marks
    doc.ActiveWindow.ActivePane.View.ShowAll = False
    
    If cc.Title = "Committee" Then
        If doc.SelectContentControlsByTitle("SummaryTitle").Count > 0 Then
            EditCC doc.SelectContentControlsByTitle("SummaryTitle").Item(1), cc.Range.Text
        End If
    End If
        
    If cc.Title = "Phases" Then
        Select Case UCase(cc.Range.Text)
            'Other
            Case "OTHER"
                HideBm "bmProOpp", "bmInvestObj", True
                HideBm "bmInvestObj", "bmProRecom", True
                HideBm "bmProRecom", "bmDetAct", True
                HideBm "bmDetAct", "bmUseForAdditional", True
            Case "STRATEGIC CASE"
                HideBm "bmProOpp", "bmInvestObj", False
                HideBm "bmInvestObj", "bmProRecom", False
                HideBm "bmProRecom", "bmDetAct", True
                HideBm "bmDetAct", "bmUseForAdditional", True
            'PBC
            Case "PROGRAMME BUSINESS CASE"
                HideBm "bmProOpp", "bmInvestObj", False
                HideBm "bmInvestObj", "bmProRecom", False
                HideBm "bmProRecom", "bmDetAct", False
                HideBm "bmDetAct", "bmUseForAdditional", True
                doc.Bookmarks("bmProRecom").Range.Tables(1).Rows(1).Cells(2).Range.Text = "Programmes considered and recommended programme"
            'IBC and DBC
            Case "INDICATIVE BUSINESS CASE", "DETAILED BUSINESS CASE"
                HideBm "bmProOpp", "bmInvestObj", False
                HideBm "bmInvestObj", "bmProRecom", False
                HideBm "bmProRecom", "bmDetAct", False
                HideBm "bmDetAct", "bmUseForAdditional", False
                doc.Bookmarks("bmProRecom").Range.Tables(1).Rows(1).Cells(2).Range.Text = "Recommended programme"
            Case Else
                HideBm "bmProOpp", "bmInvestObj", False
                HideBm "bmInvestObj", "bmProRecom", False
                HideBm "bmProRecom", "bmDetAct", False
                HideBm "bmDetAct", "bmUseForAdditional", False
                doc.Bookmarks("bmProRecom").Range.Tables(1).Rows(1).Cells(2).Range.Text = "Programmes considered and recommended programme"
        End Select
    End If  'if cc.Title="Phases"
End Sub

Function EditCC(ByRef cc As ContentControl, str As String)
    Dim blLockEdit As Boolean
    blLockEdit = cc.LockContents
    cc.LockContents = False
    cc.Range.Text = str
    cc.LockContents = blLockEdit
End Function

Function HideBm(bmF As String, bmT As String, blHide As Boolean)
    Dim rg As Range
    Set rg = ActiveDocument.Range
    If ActiveDocument.Bookmarks.Exists(bmF) And ActiveDocument.Bookmarks.Exists(bmT) Then
        rg.SetRange ActiveDocument.Bookmarks(bmF).Start, ActiveDocument.Bookmarks(bmT).Start
    Else
        Exit Function
    End If
    rg.Font.Hidden = blHide
    Dim pg As Paragraph
    For Each pg In rg.Paragraphs
        If InStr(pg.Style, "body numbered") > 0 Then
            If blHide Then
                pg.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
            Else
                pg.Range.ListFormat.ApplyListTemplate ActiveDocument.Bookmarks("bmListStyle").Range.ListFormat.ListTemplate, True
            End If
        End If
    Next pg
End Function


Sub setCC()
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If cc.Title = "Date" Then
            cc.SetPlaceholderText , , "Choose Date"
        End If
    Next cc
End Sub
