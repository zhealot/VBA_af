Attribute VB_Name = "Allfields"
Public strActiveUser As String
Public Const FixString = "PCC #"

Public Sub jk()
FilePaths.Autoexec
Dim WebNames() As String
Dim i As Integer

WebNames = IniOP.LoadIniSectionKeysArray("Links", strWebIni)
For i = LBound(WebNames) To UBound(WebNames)
    Debug.Print WebNames(i)
Next i
End Sub


'**************************************
' Replaces the contents of a bookmark
'**************************************
Public Sub ReplaceBookmarkText(sBookmark As String, sText As String, Optional Suppress As Boolean = True)
    Dim StartPos
    Dim EndPos
    
    If ActiveDocument.Bookmarks.Exists(sBookmark) Then
        Dim BMRange As Range
        Set BMRange = ActiveDocument.Bookmarks(sBookmark).Range
        BMRange.Text = sText
        ActiveDocument.Bookmarks.Add sBookmark, BMRange
    ElseIf Not Suppress Then
        MsgBox "Bookmark does not exist", vbCritical + vbOKOnly
        Exit Sub
    End If
    
End Sub

Public Function GetBookmarkText(sBookmark As String, Optional Suppress As Boolean = True) As String
    
    If ActiveDocument.Bookmarks.Exists(sBookmark) Then
        Dim BMRange As Range
        Set BMRange = ActiveDocument.Bookmarks(sBookmark).Range
        GetBookmarkText = BMRange.Text
    ElseIf Not Suppress Then
        MsgBox "Bookmark does not exist", vbCritical + vbOKOnly
        Exit Function
    End If
    
End Function


Function FileExists(fname) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(fname)
End Function

Function FolderExists(fname) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(fname)
End Function

Sub ThrowFatalError(strError As String)
    ShowError strError
    End
End Sub
Sub ShowError(strError As String)
    MsgBox strError & vbCr & vbCr & _
            "If the problem persists contact IT Support", _
                vbCritical + vbOKOnly
End Sub

Public Function PCC_Footer()
'insert Document Number by fields code from DocumentProperty->Comments
'tao@allfields.co.nz, 28/Jan/2016
    Application.ScreenUpdating = False
    Dim RgPrimary As Range
    Dim RgFirst As Range
    Dim FtPrimary As HeaderFooter
    Dim FtFirst As HeaderFooter
    Dim Doc As Document
    Dim RgTmp As Range
    Dim sc As Section
    Set Doc = ActiveDocument
    
    'update fields
    For Each sc In Doc.Sections
        If sc.Footers(wdHeaderFooterEvenPages).Exists Then
            sc.Footers(wdHeaderFooterEvenPages).Range.Fields.Update
        End If
        If sc.Footers(wdHeaderFooterFirstPage).Exists Then
            sc.Footers(wdHeaderFooterFirstPage).Range.Fields.Update
        End If
        If sc.Footers(wdHeaderFooterPrimary).Exists Then
            sc.Footers(wdHeaderFooterPrimary).Range.Fields.Update
        End If
    Next sc
    
    'do nothing if no comments
    If Trim(Doc.BuiltInDocumentProperties(wdPropertyComments).Value) = "" Then
        Exit Function
    End If
    
    
    'insert "PCC #<document Number><tab>" into first page's footer while keeping existing footer information
    Set FtFirst = Doc.Sections(1).Footers(wdHeaderFooterFirstPage)
    If Not Left(Trim(FtFirst.Range.Text), Len(FixString)) = FixString Then
        Set RgTmp = FtFirst.Range
        RgTmp.Collapse wdCollapseStart
        RgTmp.InsertBefore vbTab
        Set RgTmp = FtFirst.Range
        RgTmp.Collapse wdCollapseStart
        Doc.Fields.Add RgTmp, wdFieldComments
        Set RgTmp = FtFirst.Range
        RgTmp.Collapse wdCollapseStart
        FtFirst.Range.InsertBefore FixString
    End If
    
    Set FtPrimary = Doc.Sections(1).Footers(wdHeaderFooterPrimary)
    If Not Left(Trim(FtPrimary.Range.Text), Len(FixString)) = FixString Then
        SetPrimaryFooter FtPrimary, Doc
    End If
        
    If Doc.Sections.count > 1 Then
        Dim i As Integer
        For i = 2 To Doc.Sections.count Step 1
            If Doc.Sections(i).Footers(wdHeaderFooterEvenPages).Exists And Left(Trim(Doc.Sections(i).Footers(wdHeaderFooterEvenPages).Range.Text), Len(FixString)) <> FixString Then
                SetPrimaryFooter Doc.Sections(i).Footers(wdHeaderFooterEvenPages), Doc
            End If
            If Doc.Sections(i).Footers(wdHeaderFooterFirstPage).Exists And Left(Trim(Doc.Sections(i).Footers(wdHeaderFooterFirstPage).Range.Text), Len(FixString)) <> FixString Then
                SetPrimaryFooter Doc.Sections(i).Footers(wdHeaderFooterFirstPage), Doc
            End If
            If Doc.Sections(i).Footers(wdHeaderFooterPrimary).Exists And Left(Trim(Doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Text), Len(FixString)) <> FixString Then
                SetPrimaryFooter Doc.Sections(i).Footers(wdHeaderFooterPrimary), Doc
            End If
        Next
    End If
    
    Application.ScreenUpdating = True
End Function

Public Function SetPrimaryFooter(FooterPrimary As HeaderFooter, Doc As Document)
    FooterPrimary.Range.Text = ""
    FooterPrimary.Range.Collapse wdCollapseStart
    Doc.Fields.Add FooterPrimary.Range, wdFieldComments
    FooterPrimary.Range.Collapse wdCollapseStart
    FooterPrimary.Range.InsertBefore FixString
End Function
