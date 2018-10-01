Attribute VB_Name = "Module1"
'----------------------------------   -----------------------------------------
' Developed for NEO (Ergo)
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             September 2018
' Description:      populate job title etc, fix blank page issue.
'------------------------------------------------------------------------------

Public Const DateFormat = "dd-mm-yyyy"

Sub ShowForm(control As IRibbonControl)
    FormInit False
    UserForm1.Show
End Sub

Function FormInit(IsNew As Boolean)
    If IsNew Then
        UserForm1.tbJobTitle.Text = ""
        UserForm1.tbAuthor.Text = ""
        UserForm1.tbClient.Text = ""
        UserForm1.tbDocumentNumber.Text = ""
        UserForm1.tbDate.Text = Format(Date, DateFormat)
        UserForm1.tbRevision.Text = ""
        UserForm1.tbTitle.Text = ""
        'UserForm1.tbProject.Text = ""
        UserForm1.cbRevision.Value = True
        UserForm1.cbCover.Value = True
        UserForm1.cbCoverImage.Enabled = True
        UserForm1.cbCoverImage.Caption = "Click to select cover image"
        UserForm1.cbPDF.Value = False
    Else
        UserForm1.tbJobTitle.Text = ActiveDocument.CustomDocumentProperties("JobTitle").Value
        UserForm1.tbTitle.Text = ActiveDocument.BuiltInDocumentProperties("Title").Value
        UserForm1.tbAuthor.Text = ActiveDocument.CustomDocumentProperties("Author").Value
        UserForm1.tbClient.Text = ActiveDocument.CustomDocumentProperties("Client").Value
        UserForm1.tbDocumentNumber.Text = ActiveDocument.CustomDocumentProperties("DocumentNumber").Value
        UserForm1.tbDate.Text = ActiveDocument.CustomDocumentProperties("Date").Value
        UserForm1.tbRevision.Text = GetRevision(False)
        'UserForm1.tbProject.Text = ActiveDocument.CustomDocumentProperties("Project").Value
        UserForm1.cbRevision.Value = False
        UserForm1.tbRevision.Enabled = False
        UserForm1.Label4.Enabled = False
        UserForm1.cbCover.Value = False
        UserForm1.cbCoverImage.Enabled = False
        UserForm1.cbCoverImage.Caption = ""
        UserForm1.cbPDF.Value = IIf(Trim(ActiveDocument.CustomDocumentProperties("SecPageNo").Value) = "", False, True)
    End If
End Function

Function CheckTB(tb As TextBox, s As String) As Boolean
    If Len(Trim(tb.Text)) = 0 Then
        tb.SetFocus
        MsgBox "Please enter the " & s
        CheckTB = False
    Else
        CheckTB = True
    End If
End Function

Function UpdateCC()
    Application.ScreenUpdating = False
    SetCC "ccJobTitle", ActiveDocument.CustomDocumentProperties("JobTitle").Value
    SetCC "ccTitle", ActiveDocument.BuiltInDocumentProperties("Title").Value
    SetCC "ccClient", ActiveDocument.CustomDocumentProperties("Client").Value
    SetCC "ccDocumentNumber", ActiveDocument.CustomDocumentProperties("DocumentNumber").Value
    SetCC "ccDate", ActiveDocument.CustomDocumentProperties("Date").Value
    If UserForm1.cbRevision.Value Then
        SetCC "ccRevision", ActiveDocument.CustomDocumentProperties("Revision").Value
    End If
    SetCC "ccProject", ActiveDocument.CustomDocumentProperties("Project").Value
    Application.ScreenUpdating = True
End Function

Function SetCC(cctag As String, vl As String)
    Dim cc As ContentControl
    If ActiveDocument.SelectContentControlsByTag(cctag).Count > 0 Then
        For Each cc In ActiveDocument.SelectContentControlsByTag(cctag)
            cc.Range.Text = vl
        Next cc
    End If
End Function

'get revision value from table in page 2
Function GetRevision(add As Boolean) As String
    Dim tb As Table
    GetRevision = ""
    If ActiveDocument.Bookmarks.Exists("bmRevision") And ActiveDocument.Bookmarks.Exists("bmRD") Then
        Set tb = ActiveDocument.Bookmarks("bmRevision").Range.Tables(1)
        Dim i As Integer
        Dim RDIndex As Integer
        RDIndex = ActiveDocument.Bookmarks("bmRD").Range.Cells(1).RowIndex
        'first part
        For i = 3 To RDIndex - 2
            If tb.Rows(i).Cells.Count = 6 And Len(Trim(tb.Rows(i).Cells(1).Range.Text)) = 2 Then
                If i = 3 Then   'no revision record
                    GetRevision = ""
                Else
                    GetRevision = tb.Rows(i - 1).Cells(1).Range.Text
                    GetRevision = Left(GetRevision, Len(GetRevision) - 2)
                End If
                'add new entiry
                If add Then
                    tb.Rows(i).Cells(1).Range.Text = ActiveDocument.CustomDocumentProperties("Revision").Value 'UserForm1.tbRevision.text
                    tb.Rows(i).Cells(2).Range.Text = ActiveDocument.CustomDocumentProperties("Date").Value  'Format(UserForm1.tbDate.text, DateFormat)
                    tb.Rows(i).Cells(3).Range.Text = ActiveDocument.CustomDocumentProperties("Author").Value  'UserForm1.tbAuthor.text
                End If
                Exit For
            End If
        Next
        'second part
        For i = RDIndex + 2 To tb.Rows.Count - 1
            If Len(Trim(tb.Rows(i).Cells(1).Range.Text)) = 2 Then
                If add Then
                    tb.Rows(i).Cells(1).Range.Text = UserForm1.tbRevision.Text
                End If
                Exit Function
            End If
        Next
    End If
End Function

Function AddRevision()
    Application.ScreenUpdating = False
    Dim tb As Table
    If ActiveDocument.Bookmarks.Exists("bmRevision") And ActiveDocument.Bookmarks.Exists("bmRD") Then
        Set tb = ActiveDocument.Bookmarks("bmRevision").Range.Tables(1)
        Dim i As Integer
        Dim iRw As Integer 'row index to add new entity
        iRw = 0
        Dim RDRow As Integer 'Revision Details row index
        RDRow = ActiveDocument.Bookmarks("bmRD").Range.Cells(1).RowIndex
        '"Document history and status" part
        For i = 3 To RDRow - 2
            If tb.Rows(i).Cells.Count = 5 And Len(Trim(tb.Rows(i).Cells(1).Range.Text)) = 2 Then
                iRw = i
                Exit For
            End If
        Next i
        If iRw = 0 Then 'need to add new row
            tb.Rows(RDRow - 2).Select
            Selection.Collapse wdCollapseStart
            Selection.InsertRowsBelow
            Selection.Collapse wdCollapseStart
            iRw = Selection.Rows(1).Index
        End If
        If tb.Rows(iRw).Cells.Count = 5 Then
            tb.Rows(iRw).Cells(1).Range.Text = ActiveDocument.CustomDocumentProperties("Revision").Value 'UserForm1.tbRevision.text
            tb.Rows(iRw).Cells(2).Range.Text = ActiveDocument.CustomDocumentProperties("Date").Value  'Format(UserForm1.tbDate.text, DateFormat)
            tb.Rows(iRw).Cells(3).Range.Text = ActiveDocument.CustomDocumentProperties("Author").Value  'UserForm1.tbAuthor.text
        End If
        '"Revision Details" part
        iRw = 0
        For i = ActiveDocument.Bookmarks("bmRD").Range.Cells(1).RowIndex + 2 To tb.Rows.Count - 1
            If tb.Rows(i).Cells.Count = 2 And Len(Trim(tb.Rows(i).Cells(1).Range.Text)) = 2 Then
                iRw = i
                Exit For
            End If
        Next i
        If iRw = 0 Then 'need to add new row
            tb.Rows(tb.Rows.Count - 1).Select
            Selection.Collapse wdCollapseStart
            Selection.InsertRowsBelow
            Selection.Collapse wdCollapseStart
            iRw = Selection.Rows(1).Index
        End If
        If tb.Rows(iRw).Cells.Count = 2 Then
            tb.Rows(iRw).Cells(1).Range.Text = ActiveDocument.CustomDocumentProperties("Revision").Value 'UserForm1.tbRevision.text
        End If
    Else
        MsgBox "bookmark not found: bmRevision/bmRD"
    End If
    Application.ScreenUpdating = True
End Function
Sub defa()
    Dim tb As Table
    Set tb = ThisDocument.Bookmarks("bmRD").Range.Tables(1)
    tb.Rows(6).Select
    Selection.Collapse wdCollapseStart
    Selection.InsertRowsBelow
    Selection.Collapse wdCollapseStart
    Debug.Print Selection.Rows(1).Index
End Sub


Sub test()
    Dim ary() As String
    ary = Split(",sd,vw,234,xcv,234ds,xcert", ",")
    Debug.Print UBound(ary)
    
End Sub

