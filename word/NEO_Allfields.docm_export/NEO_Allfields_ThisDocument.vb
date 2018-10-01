VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'----------------------------------   -----------------------------------------
' Developed for NEO (Ergo)
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             September 2018
' Description:      populate job title etc, fix blank page issue.
'------------------------------------------------------------------------------

Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    Dim cc As ContentControl
    Dim rg As Range
    Dim i As Integer
    Dim j As Integer
    
    For Each rg In ActiveDocument.StoryRanges
        For Each cc In rg.ContentControls
            If ContentControl.Tag <> "" And cc.Tag = ContentControl.Tag Then
                If cc.ID <> ContentControl.ID Then
                    cc.Range.Text = ContentControl.Range.Text
                    Select Case ContentControl.Tag
                    Case "ccJobTitle"
                        ActiveDocument.CustomDocumentProperties("JobTitle").Value = ContentControl.Range.Text
                    Case "ccTitle"
                        ActiveDocument.BuiltInDocumentProperties("Title").Value = ContentControl.Range.Text
                    Case "ccClient"
                        ActiveDocument.CustomDocumentProperties("Client").Value = ContentControl.Range.Text
                    Case "ccDocumentNumber"
                        ActiveDocument.CustomDocumentProperties("DocumentNumber").Value = ContentControl.Range.Text
                    Case "ccDate"
                        ActiveDocument.CustomDocumentProperties("Date").Value = ContentControl.Range.Text
                    Case "ccRevision"
                        ActiveDocument.CustomDocumentProperties("Revision").Value = ContentControl.Range.Text
                    Case "ccProject"
                        ActiveDocument.CustomDocumentProperties("Project").Value = ContentControl.Range.Text
                    Case Else
                    End Select
                End If
            End If
        Next cc
    Next rg
    
    For i = 1 To ActiveDocument.Sections.Count
        For j = 1 To 3
            'If ActiveDocument.Sections(i).Footers(j).Exists Then
                Set rg = ActiveDocument.Sections(i).Footers(j).Range
                If rg.ContentControls.Count > 0 Then
                    For Each cc In rg.ContentControls
                        If ContentControl.Tag <> "" And cc.Tag = ContentControl.Tag Then
                            If cc.ID <> ContentControl.ID Then
                                cc.Range.Text = ContentControl.Range.Text
                            End If
                        End If
                    Next cc
                End If
            'End If
        Next j
    Next
End Sub

Private Sub Document_New()
    UserForm1.Show
End Sub

Private Sub Document_Open()
    If ThisDocument.CustomDocumentProperties("IsNew").Value = "Yes" Then
        UserForm1.Show
    End If
End Sub
