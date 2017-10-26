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
    Dim doc As Document
    Dim RgTmp As Range
    Dim sc As Section
    Set doc = ActiveDocument
    
    'update footer fields
    For Each sc In doc.Sections
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
    
    'update header fields
    For Each sc In doc.Sections
        If sc.Headers(wdHeaderFooterEvenPages).Exists Then
            sc.Headers(wdHeaderFooterEvenPages).Range.Fields.Update
        End If
        If sc.Headers(wdHeaderFooterFirstPage).Exists Then
            sc.Headers(wdHeaderFooterFirstPage).Range.Fields.Update
        End If
        If sc.Headers(wdHeaderFooterPrimary).Exists Then
            sc.Headers(wdHeaderFooterPrimary).Range.Fields.Update
        End If
    Next sc
    
    
    'do nothing if no comments
    If Trim(doc.BuiltInDocumentProperties(wdPropertyComments).Value) = "" Then
        Exit Function
    End If
        
    'insert "PCC #<document Number><tab>" into first page's footer while keeping existing footer information
'    Set FtFirst = doc.Sections(1).Footers(wdHeaderFooterFirstPage)
'    If Not Left(Trim(FtFirst.Range.Text), Len(FixString)) = FixString Then
'        Set RgTmp = FtFirst.Range
'        RgTmp.Collapse wdCollapseStart
'        RgTmp.InsertBefore vbTab
'        Set RgTmp = FtFirst.Range
'        RgTmp.Collapse wdCollapseStart
'        doc.Fields.Add RgTmp, wdFieldComments
'        Set RgTmp = FtFirst.Range
'        RgTmp.Collapse wdCollapseStart
'        FtFirst.Range.InsertBefore FixString
'    End If
'
'    Set FtPrimary = doc.Sections(1).Footers(wdHeaderFooterPrimary)
'    If Not Left(Trim(FtPrimary.Range.Text), Len(FixString)) = FixString Then
'        SetPrimaryFooter FtPrimary, doc
'    End If
'
'    If doc.Sections.count > 1 Then
'        Dim i As Integer
'        For i = 2 To doc.Sections.count Step 1
'            If doc.Sections(i).Footers(wdHeaderFooterEvenPages).Exists And Left(Trim(doc.Sections(i).Footers(wdHeaderFooterEvenPages).Range.Text), Len(FixString)) <> FixString Then
'                SetPrimaryFooter doc.Sections(i).Footers(wdHeaderFooterEvenPages), doc
'            End If
'            If doc.Sections(i).Footers(wdHeaderFooterFirstPage).Exists And Left(Trim(doc.Sections(i).Footers(wdHeaderFooterFirstPage).Range.Text), Len(FixString)) <> FixString Then
'                SetPrimaryFooter doc.Sections(i).Footers(wdHeaderFooterFirstPage), doc
'            End If
'            If doc.Sections(i).Footers(wdHeaderFooterPrimary).Exists And Left(Trim(doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Text), Len(FixString)) <> FixString Then
'                SetPrimaryFooter doc.Sections(i).Footers(wdHeaderFooterPrimary), doc
'            End If
'        Next
'    End If
    Application.ScreenUpdating = True
End Function

Public Function SetPrimaryFooter(FooterPrimary As HeaderFooter, doc As Document)
    FooterPrimary.Range.Text = ""
    FooterPrimary.Range.Collapse wdCollapseStart
    doc.Fields.Add FooterPrimary.Range, wdFieldComments
    FooterPrimary.Range.Collapse wdCollapseStart
    FooterPrimary.Range.InsertBefore FixString
End Function


Function FillBookmark(doc As Document, bm As String, txt As String)
    If doc.Bookmarks.Exists(bm) Then
        doc.Bookmarks(bm).Range.Text = txt
    End If
End Function

Sub TestAD()
'test reading user info from AD via AD object
    Dim sUser As String
    Set objSysInfo = CreateObject("ADSystemInfo")
    sUser = objSysInfo.UserName
    Set objuser = GetObject("LDAP://" & sUser)
    
    Debug.Print sUser
    Debug.Print objuser.DisplayName             'user name
    Debug.Print objuser.Title                   'title
    Debug.Print objuser.department              'department
    Debug.Print objuser.telephoneNumber         'phone
    Debug.Print objuser.facsimileTelephoneNumber 'fax
    Debug.Print objuser.mobile                  'mobile
    Debug.Print objuser.mail                    'email

End Sub
