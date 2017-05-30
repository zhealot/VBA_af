VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDocProp 
   Caption         =   "Report Properties"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8640
   OleObjectBlob   =   "MMNZ Report Template_1.11 April 2016_Allfields_frmDocProp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDocProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdUpdate_Click()

Dim strClient, strReportDate, strReportTitle, strHeaderSum, strReportNo, strRepSecurity, strAuthors, strChecked As String

'   Check Client Organisation

    strClient = Trim(frmDocProp.txtClient.Value) 'Get value from the form
    If strClient = "" Then
        MsgBox "You must specify the Client Organisation", vbCritical, "Update Report Properties"
        frmDocProp.txtClient.SetFocus 'Position cursor
        Exit Sub
    Else
        ActiveDocument.CustomDocumentProperties("Client") = strClient
    End If
    
'   Check Report Date

    strReportDate = Trim(frmDocProp.txtReportDate.Value) 'Get value from the form
    If strReportDate = "" Then
        MsgBox "You must specify the Report Date", vbCritical, "Update Report Properties"
        frmDocProp.txtReportDate.SetFocus 'Position cursor
        Exit Sub
    Else
        ActiveDocument.CustomDocumentProperties("Report Date") = strReportDate
    End If
    
'   Check Report Title

    strReportTitle = Trim(frmDocProp.txtReportTitle.Value) 'Get value from the form
    If strReportTitle = "" Then
        MsgBox "You must specify the Report Title of this document.", vbCritical, "Update Report Properties"
        frmDocProp.txtReportTitle.SetFocus 'Position cursor
        Exit Sub
    Else
        ActiveDocument.CustomDocumentProperties("Report Title") = strReportTitle
    End If
    
'   Check Header Summary Name
    
    strHeaderSum = Trim(frmDocProp.txtHeaderSummary.Value) 'Get value from the form
    If strHeaderSum = "" Then
        MsgBox "You must specify the Header Summary Name", vbCritical, "Update Report Properties"
        frmDocProp.txtHeaderSummary.SetFocus 'Position cursor in the Number box
        Exit Sub
    Else
        ActiveDocument.CustomDocumentProperties("Header Summary") = strHeaderSum
    End If

'   Check Report Number
    
    strReportNo = Trim(frmDocProp.txtReportNo.Value) 'Get value from the form
    If strReportNo = "" Then
        MsgBox "You must specify the Report Number of this document.", vbCritical, "Update Report Properties"
        frmDocProp.txtReportNo.SetFocus 'Position cursor
        Exit Sub
    Else
        ActiveDocument.CustomDocumentProperties("Report No") = strReportNo
    End If
    
'   Check Issue

    If frmDocProp.cboIssue.ListIndex <> -1 Then
        ActiveDocument.CustomDocumentProperties("Issue") = Trim(frmDocProp.cboIssue.Value)
    Else
        MsgBox "Please select the required value from the Issue drop list.", vbCritical, "Update Report Properties"
        frmDocProp.cboIssue.SetFocus
        Exit Sub
    End If
    
'   Check Security
    
    strRepSecurity = Trim(frmDocProp.cboSecurity.Value)
    ActiveDocument.CustomDocumentProperties("Report Security") = strRepSecurity
    
'   Check Authors

    strAuthors = Trim(frmDocProp.txtAuthors.Value) 'Get value from the form
    If strAuthors = "" Then
        MsgBox "You must specify the Author/s of this document.", vbCritical, "Update Report Properties"
        frmDocProp.txtAuthors.SetFocus 'Position cursor in the Number box
        Exit Sub
    Else
        ActiveDocument.CustomDocumentProperties("Authors") = strAuthors
    End If

'   Checked by

    strChecked = Trim(frmDocProp.txtChecked.Value) 'Get value from the form
    If strChecked = "" Then
        MsgBox "You must specify who this document is checked by.", vbCritical, "Update Report Properties"
        frmDocProp.txtChecked.SetFocus 'Position cursor in the Number box
        Exit Sub
    Else
        ActiveDocument.CustomDocumentProperties("Checked") = strChecked
    End If
    
'   Select the whole document and update the fields
    
    Application.ScreenUpdating = False
    Selection.WholeStory
    Selection.Fields.Update
    Selection.HomeKey Unit:=wdStory
    
'   Use Print preview to update the fields in the headers/footers
    
    ActiveDocument.PrintPreview
    ActiveDocument.ClosePrintPreview
    Application.ScreenUpdating = True
    frmDocProp.Hide  'Hide the form
    Application.WindowState = wdWindowStateMaximize

End Sub


Private Sub CommandButton1_Click()
'Sub exportComments()
' Exports comments from a MS Word document to Excel and associates them with the heading paragraphs
' they are included in. Useful for outline numbered section, i.e. 3.2.1.5....
' Thanks to Graham Mayor, http://answers.microsoft.com/en-us/office/forum/office_2007-customize/export-word-review-comments-in-excel/54818c46-b7d2-416c-a4e3-3131ab68809c
' and Wade Tai, http://msdn.microsoft.com/en-us/library/aa140225(v=office.10).aspx
' Need to set a VBA reference to "Microsoft Excel 14.0 Object Library"

Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
Dim i As Integer, HeadingRow As Integer
Dim objPara As Paragraph
Dim objComment As Comment
Dim strSection As String
Dim strTemp
Dim myRange As Range

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True
Set xlWB = xlApp.Workbooks.Add 'create a new workbook
With xlWB.Worksheets(1)
' Create Heading
    HeadingRow = 1
    .Cells(HeadingRow, 1).Formula = "Comment"
    .Cells(HeadingRow, 2).Formula = "Page"
    .Cells(HeadingRow, 3).Formula = "Paragraph"
    .Cells(HeadingRow, 4).Formula = "Comment"
    .Cells(HeadingRow, 5).Formula = "Reviewer"
    .Cells(HeadingRow, 6).Formula = "Date"
    .Cells(HeadingRow, 7).Formula = "Action"
    .Cells(HeadingRow, 8).Formula = "Action Addressee"
    
    strSection = "preamble" 'all sections before "1." will be labeled as "preamble"
    strTemp = "preamble"
    If ActiveDocument.Comments.Count = 0 Then
        MsgBox ("No comments")
        Exit Sub
    End If
    
    For i = 1 To ActiveDocument.Comments.Count
        Set myRange = ActiveDocument.Comments(i).Scope
        strSection = ParentLevel(myRange.Paragraphs(1)) ' find the section heading for this comment
        'MsgBox strSection
        .Cells(i + HeadingRow, 1).Formula = ActiveDocument.Comments(i).Index
        .Cells(i + HeadingRow, 2).Formula = ActiveDocument.Comments(i).Reference.Information(wdActiveEndAdjustedPageNumber)
        .Cells(i + HeadingRow, 3).Value = strSection
        .Cells(i + HeadingRow, 4).Formula = ActiveDocument.Comments(i).Range
        .Cells(i + HeadingRow, 5).Formula = ActiveDocument.Comments(i).Initial
        .Cells(i + HeadingRow, 6).Formula = Format(ActiveDocument.Comments(i).Date, "dd/MM/yyyy")
        .Cells(i + HeadingRow, 7).Formula = ActiveDocument.Comments(i).Range.ListFormat.ListString
    Next i
End With
Set xlWB = Nothing
Set xlApp = Nothing

End Sub

'Function ParentLevel(Para As Word.Paragraph) As String
'From Tony Jollans
' Finds the first outlined numbered paragraph above the given paragraph object
    'Dim ParaAbove As Word.Paragraph
    'Set ParaAbove = Para
    'sStyle = Para.Range.ParagraphStyle
    'sStyle = ParaAbove.Range.ParagraphStyle
    'sStyle = Left(sStyle, 4)
   ' If sStyle = "Head" Then
    '    GoTo Skip
    'End If
    'Do Until ParaAbove.OutlineLevel = Para.OutlineLevel
    '    Set ParaAbove = ParaAbove.Previous
   ' Loop
'Skip:
   ' strTitle = ParaAbove.Range.Text
   'strTitle = Left(strTitle, Len(strTitle) - 1)
    'ParentLevel = ParaAbove.Range.ListFormat.ListString & " " & strTitle
'End Function


Private Sub UserForm_Initialize()

'   Populate Default Values
    txtReportDate.Text = Date
    txtCompany.Text = "Marico Marine"
    txtAuthors.Text = ActiveDocument.BuiltInDocumentProperties("Author")
    
'   Populate Issue Values
    cboIssue.List = Array("Vs1", "Vs2", "Vs3", "Vs4", "Vs5", "Vs6", "Vs7", "Vs8", "Draft A", "Draft B", "Draft C", "Draft D", "Draft E", "Draft F", "Draft G" _
    , "Draft H", "Draft I", "Draft J", "Draft K", "Draft L", "Draft M", "01", "02", "03", "04", "05", "06" _
    , "07", "08")
    
'   Populate Security Values
'###add: "Unrestricted", tao@allfields.co.nz 28/10/2015
    cboSecurity.List = Array("", "Restricted", "Unrestricted", "Commercial-in-Confidence")
    


    
End Sub

Private Sub txtClient_AfterUpdate()

    lblClientHeader.Caption = txtClient.Text
    
End Sub



