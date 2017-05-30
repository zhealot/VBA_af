VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'2)  Issuing Authority section
'- text from original document (first part) not copied over
'
'4)  Blue/violet highlighting
'I noticed that this is where the text in the original document is highlighted as "editable" and came across into an editable area


Private Sub CommandButton1_Click()
    Dim docTpl As Document
    Dim docOri As Document
    Dim sFilePath As String
    Dim doc As Document
    Dim rgTpl As Range
    Dim rgOri As Range
    Dim cc As ContentControl
    Dim ccTpl As ContentControl
    Dim ccOri As ContentControl
    Dim sFile As String
    Dim tmpStr As String
    Dim TypeOfInstrument As String
    Dim mPty As MetaProperty
    Dim strAry(2) As String
    Dim errMsg As String
    Dim sTplPath As String: sTplPath = "https://piritahi.cohesion.net.nz/Sites/RG/TemplateDocs/"
    Dim iFileOpened As Integer: iFileOpened = 0
    Dim iFileExported As Integer: iFileExported = 0
    Dim ccLocked As Boolean
    Dim iRgStart As Integer
    Dim iRgEnd As Integer
    Dim tbOri As Table
    Dim tbTpl As Table
    Dim iCounter As Integer
    Dim jCounter As Integer
        
    'check if logged into SP
    Dim oHttpRequest
    Err.Clear
    Set oHttpRequest = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    With oHttpRequest
        .Open "GET", sTplPath, False
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Pragma", "no-cache"
        .send
    End With
    If InStr(Err.Description, "Access is denied") > 0 Then 'If Err.Number <> 0 Or oHttpRequest.Status <> 200 Then   'HTTP code 200 means OK
        MsgBox "Access denied: " & vbNewLine & sTplPath
        Exit Sub
    End If
    On Error GoTo 0
    
    Set doc = ThisDocument
    
    If Dir(doc.Path & "\Processed", vbDirectory) = "" Then
        MkDir doc.Path & "\Processed\"
    End If
    
    Err.Clear
    sFilePath = doc.Path & "\"
    sFile = Dir(sFilePath & "*.docx")

    'for each file in same folder with this Migratio Tool
    Do While sFile <> ""
        On Error GoTo OpenFail
        sFile = sFilePath & sFile
        errMsg = sFile
        Set docOri = Documents.Open(sFile, ReadOnly:=True, Visible:=False)
        iFileOpened = iFileOpened + 1
        Debug.Print "Ori docx: " & docOri.Path & "\" & docOri.Name
        Log "Open docx: " & docOri.Path & "\" & docOri.Name
        
        'delete restricted edit areas to eliminate background colour, tao@allfields.co.nz 3/12/2015
        On Error Resume Next
        Dim pr As Paragraph
        Dim iEditor As Integer
        For Each pr In docOri.Paragraphs
            If pr.Range.Editors.Count > 0 Then
                For iEditor = 1 To pr.Range.Editors.Count
                    pr.Range.Editors(iEditor).Delete
                Next iEditor
            End If
        Next pr
        
        On Error GoTo NoCC
        errMsg = "TypeOfInstrument"
        TypeOfInstrument = docOri.SelectContentControlsByTitle("TypeOfInstrument").Item(1).Range.Text
        
        On Error GoTo OpenFail
        If ToITemplate(TypeOfInstrument) = "" Then
            errMsg = "No template file found: " & TypeOfInstrument
            GoTo OpenFail
        End If
        sFile = sTplPath & ToITemplate(TypeOfInstrument)
        errMsg = sFile
        Set docTpl = Documents.Open(sFile, ReadOnly:=True, Visible:=False) ', Format:=wdOpenFormatXMLDocumentMacroEnabled)
        '###Set docTpl = Documents.Open("C:\Users\tao\Box Sync\2. Staff Related Activities\Tao's\RGP\MigrationTool\Guidance\MPI_Guidance_Template.docx", ReadOnly:=True, Visible:=True)
        Debug.Print "Template doc: " & docTpl.Path & "\" & docTpl.Name
        Log "Open Template doc: " & docTpl.Path & "\" & docTpl.Name
        docTpl.Windows(1).View.Type = wdPrintView

        'set DIP title
        On Error Resume Next
        tmpStr = ""
        tmpStr = docOri.SelectContentControlsByTitle("TitleOfDocument").Item(1).Range.Text
        If Err.Number <> 0 Then
            Log "!!!!No CC: TitleOfDocument"
        End If
        Err.Clear
        docTpl.BuiltInDocumentProperties("Title").Value = tmpStr
        If Err.Number <> 0 Then
            Log "!!!!!Set title failed: "
        Else
            Log ("Set DIP title" & vbTab & tmpStr)
        End If
        'set title cc in case not populated by DIP
        
        For Each cc In docTpl.SelectContentControlsByTitle("Title")
            ccLocked = cc.LockContents
            cc.LockContents = False
            'make front page title cc richtext
            If cc.Range.Information(wdActiveEndAdjustedPageNumber) = 1 And cc.Type = wdContentControlText Then
                cc.Type = wdContentControlRichText
                Log ("Convert front page title CC into rich text CC")
            End If
            cc.Range.Text = tmpStr
            cc.LockContents = ccLocked
        Next cc
        
        Err.Clear
        tmpStr = docOri.SelectContentControlsByTitle("SubtitleOfDocument").Item(1).Range.Text
        If Err.Number <> 0 Then
            Log "!!!!!No CC: SubtitleOfDocument"
        End If
        Err.Clear
        docTpl.ContentTypeProperties("Subtitle").Value = tmpStr
        If Err.Number <> 0 Then
            Log "!!!!!Set subtitle failed"
        Else
            Log "Set subtitle" & vbTab & tmpStr
        End If
       
        'set Document Type
        For Each mPty In docTpl.ContentTypeProperties
            If mPty.Name = "Document Type" Then
                PopulateArray strAry, TypeOfInstrument
                If strAry(0) = "" Or strAry(1) = "" Then
                    errMsg = "Document Type Value for " & TypeOfInstrument & " not found"
                    GoTo NoCC
                End If
                mPty.Value = strAry
                Log ("Set DIP Document Type:" & vbTab & TypeOfInstrument)
                Exit For
            End If
        Next mPty
        
        'has Revocation/RelatedRequirements
        Dim sCCinFrontPage As String
        Select Case TypeOfInstrument
            Case "Guidance Document"
                sCCinFrontPage = "RelatedRequirements"
            Case Else
                sCCinFrontPage = "Revocation"
        End Select
        If docOri.SelectContentControlsByTitle(sCCinFrontPage).Count > 0 Then
            Set rgOri = docOri.SelectContentControlsByTitle("Revocation").Item(1).Range
            ToggleContentControlPresentOrAbsent _
                doc:=docTpl _
                , tagOfContentControl:="Revocation" _
                , isControlPresentOrAbsent:=contentControlExists("Revocation") _
                , tagOfContentControl2insertAfterOrInside:="Commencement" _
                , insertBeforeInsideOrAfter:=jcAfter
                Log ("Insert Revocation cc")
            If docTpl.SelectContentControlsByTitle("Revocation").Count = 0 Then
                Log "!!!!!CC Revocation not inserted"
            Else
                Set rgTpl = docTpl.SelectContentControlsByTitle("Revocation").Item(1).Range
                If rgOri.Paragraphs.Count > 1 Then
                    rgOri.SetRange rgOri.Paragraphs(2).Range.Start, rgOri.End + 2
                    rgOri.Copy
                    rgTpl.Collapse wdCollapseEnd
                    rgTpl.PasteAndFormat wdUseDestinationStylesRecovery
                End If
            End If
            Set rgOri = Nothing
        End If
        
        'Issuing authority part
        If docOri.SelectContentControlsByTitle("IssuingAuthority").Count > 0 _
            And docTpl.SelectContentControlsByTitle("IssuingAuthority").Count > 0 Then
            Set rgOri = docOri.SelectContentControlsByTitle("IssuingAuthority").Item(1).Range
            Set rgTpl = docTpl.SelectContentControlsByTitle("IssuingAuthority").Item(1).Range
           
            With rgOri.Find
                .ClearAllFuzzyOptions
                .ClearFormatting
                .Wrap = wdFindStop
                .MatchCase = False
                .MatchWholeWord = True
                .Text = "is issued"
                .Execute
            End With
            With rgTpl.Find
                .ClearAllFuzzyOptions
                .ClearFormatting
                .Wrap = wdFindStop
                .MatchCase = False
                .MatchWholeWord = True
                .Text = "is issued"
                .Execute
            End With
            If rgOri.Find.Found And rgTpl.Find.Found Then
                rgOri.SetRange rgOri.End, docOri.SelectContentControlsByTitle("IssuingAuthority").Item(1).Range.End - 1
                rgTpl.SetRange rgTpl.End, docTpl.SelectContentControlsByTitle("IssuingAuthority").Item(1).Range.End - 1
                rgTpl.Text = ""
                rgTpl.Collapse wdCollapseStart
                rgTpl.Text = rgOri.Text
            End If
            Set rgOri = Nothing
            Set rgTpl = Nothing
        End If
                
        'Purpose for Guidance document
        If docOri.SelectContentControlsByTitle("Purpose").Count > 0 And _
        docTpl.SelectContentControlsByTitle("Purpose").Count > 0 Then
            Set ccOri = docOri.SelectContentControlsByTitle("Purpose").Item(1)
            Set ccTpl = docTpl.SelectContentControlsByTitle("Purpose").Item(1)
            If ccOri.Range.Information(wdActiveEndPageNumber) = 2 And _
            ccTpl.Range.Information(wdActiveEndPageNumber) = 2 Then
                ccLocked = ccTpl.LockContents
                ccTpl.LockContents = False
                ccTpl.Range.Text = ccOri.Range.Text
                ccTpl.LockContents = ccLocked
            End If
        End If
        
        'Document history for Guidance document
        Set rgOri = docOri.Range
        With rgOri.Find
            .ClearAllFuzzyOptions
            .ClearFormatting
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .Text = "Document history"
            .Execute
        End With
        If rgOri.Find.Found Then
            If InStr(rgOri.Style, "Heading") > 0 Then
                iRgStart = rgOri.Next(wdParagraph, 1).Start
                Set rgOri = docOri.Range
                With rgOri.Find
                    .ClearAllFuzzyOptions
                    .ClearFormatting
                    .Wrap = wdFindStop
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .Text = "Contact details"
                    .Execute
                End With
                If rgOri.Find.Found Then
                    If InStr(rgOri.Style, "Heading") > 0 Then
                        If docTpl.SelectContentControlsByTitle("History").Count > 0 And _
                        docTpl.SelectContentControlsByTitle("History").Item(1).Range.Tables.Count > 0 Then
                            rgOri.SetRange iRgStart, rgOri.Start - 1
                            Set tbTpl = docTpl.SelectContentControlsByTitle("History").Item(1).Range.Tables(1)
                            If rgOri.Tables.Count > 0 Then  'a table in document history
                                Set tbOri = rgOri.Tables(1)
                                If tbTpl.Columns.Count = tbOri.Columns.Count Then
                                    Dim blIsSame As Boolean
                                    blIsSame = True
                                    For iCounter = 1 To tbTpl.Columns.Count
                                        If Not LCase(tbOri.Cell(1, iCounter).Range.Text) = LCase(tbTpl.Cell(1, iCounter).Range.Text) Then
                                            blIsSame = False
                                            Exit For
                                        End If
                                    Next iCounter
                                    If blIsSame And tbOri.Rows.Count > 1 Then
                                        For iCounter = 2 To tbOri.Rows.Count Step 1
                                            If tbTpl.Rows.Count < iCounter Then
                                                tbTpl.Rows.Add
                                            End If
                                            For jCounter = 1 To tbOri.Columns.Count Step 1
                                                tbTpl.Cell(iCounter, jCounter).Range.Text = Replace(tbOri.Cell(iCounter, jCounter).Range.Text, Chr(13) & Chr(7), "")
                                            Next jCounter
                                        Next iCounter
                                    End If
                                End If
                            Else    'no table in Document history area, copy text into table's last "Change(s) Description" cell
                                tbTpl.Cell(2, 4).Range.Text = Trim(rgOri.Text)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        'Contact details for Guidance document
        If TypeOfInstrument = "Guidance Document" Then
            Set rgOri = docOri.Range
            With rgOri.Find 'look for 'Contact details'
                .Text = "Contact details"
                .Execute
            End With
            If rgOri.Find.Found Then
                If InStr(rgOri.Style, "Heading") > 0 Then
                    iRgStart = rgOri.Next(wdParagraph, 1).Start
                    Set rgOri = docOri.Range
                    With rgOri.Find 'look for 'Disclaimer'
                        .ClearFormatting
                        .Wrap = wdFindStop
                        .Forward = True
                        .MatchCase = False
                        .MatchWholeWord = False
                        .Format = True
                        .Text = "Disclaimer"
                        .Execute
                    End With
                    If rgOri.Find.Found And InStr(rgOri.Style, "Heading") > 0 And _
                    docTpl.SelectContentControlsByTitle("Contact").Count > 0 Then
                        rgOri.SetRange iRgStart, rgOri.Start - 2
                        docTpl.SelectContentControlsByTitle("Contact").Item(1).Range.Text = rgOri.Text
                    End If
                End If
            End If
        End If
        'RelatedRequirements for Guidance Documents
        If docOri.SelectContentControlsByTitle("RelatedRequirements").Count > 0 And _
        docTpl.SelectContentControlsByTitle("RelatedRequirements").Count > 0 Then
            Set rgOri = docOri.SelectContentControlsByTitle("RelatedRequirements").Item(1).Range
            Set rgTpl = docTpl.SelectContentControlsByTitle("RelatedRequirements").Item(1).Range
            If rgOri.Paragraphs.Count > 1 Then
                rgOri.SetRange rgOri.Paragraphs(2).Range.Start, docOri.SelectContentControlsByTitle("RelatedRequirements").Item(1).Range.End - 1
                If rgTpl.Paragraphs.Count < 2 Then
                    rgTpl.Paragraphs.Add
                End If
                rgTpl.SetRange rgTpl.Paragraphs(2).Range.Start, docTpl.SelectContentControlsByTitle("RelatedRequirements").Item(1).Range.End
                rgTpl.Text = rgOri.Text
            End If
        End If
        
        'copy body content
        Set rgOri = docOri.Range
        With rgOri.Find
            .ClearFormatting
            .Format = True
            Select Case TypeOfInstrument
                Case "Guidance Document"
                    .Style = "Heading 1"
                    .Text = "Purpose"
                Case Else
                    .Style = "Heading 1a"
                    .Text = "Introduction"
            End Select
            .Execute
            If .Found Then
                rgOri.SetRange rgOri.Start, docOri.Range.End
                rgOri.Copy
            Else
                Log "!!! Body content not found in original docx"
            End If
        End With
        Set rgTpl = docTpl.Range
        With rgTpl.Find
            .ClearFormatting
            .Format = True
            Select Case TypeOfInstrument
                Case "Guidance Document"
                    .Style = "Heading 1"
                    .Text = "Purpose"
                Case Else
                    .Style = "Heading 1a"
                    .Text = "Introduction"
            End Select
            .Execute
            If .Found Then
                rgTpl.SetRange rgTpl.Start, docTpl.Range.End
                For Each cc In docTpl.ContentControls
                    If cc.Range.Start <= rgTpl.End And cc.Range.End >= rgTpl.Start Then
                        cc.LockContentControl = False
                        cc.LockContents = False
                        cc.Delete
                    End If
                Next cc
                rgTpl.Delete
                rgTpl.Collapse wdCollapseStart
                'wdUseDestinationStylesRecovery
                rgTpl.PasteAndFormat wdListRestartNumbering
                Log ("Body content copied")
            End If
        End With
        
        'sort out paragraphs numbering
        On Error Resume Next
        Dim pg As Paragraph
        Dim rgPg As Range
        Dim Style As Style
        Dim StylePre As Style
        
        If TypeOfInstrument = "Guidance Document" Then
            For Each pg In docTpl.Paragraphs
                Set rgPg = pg.Range
                Set Style = rgPg.Style
                If Style.NameLocal = "Clause L1" And rgPg.ListParagraphs.Count > 0 Then
                    Set StylePre = pg.Previous.Range.Style
                    If InStr(StylePre.NameLocal, "Heading") > 0 Or StylePre.NameLocal = "Normal Indent" Then
                        rgPg.ListFormat.ApplyListTemplate ListTemplate:=rgPg.ListFormat.ListTemplate, ContinuePreviousList:=False, applyto:=wdlistthispointforward
                    End If
                End If
                If Style.NameLocal = "Heading 1" And rgPg.ListParagraphs.Count > 0 Then
                    rgPg.ListFormat.ApplyListTemplate ListTemplate:=rgPg.ListFormat.ListTemplate, ContinuePreviousList:=True, applyto:=wdlistthispointforward
                End If
            Next pg
        Else
            Set Style = docTpl.Styles("Heading 2a")
            For Each pg In docTpl.Paragraphs
                Set rgPg = pg.Range
                If pg.Style = "Clause L1" And rgPg.ListParagraphs.Count > 0 Then
                    If InStr(pg.Previous.Style, "Heading 1") > 0 _
                    Or InStr(pg.Previous.Style, "Heading 2") > 0 _
                    Or (InStr(pg.Previous.Style, "Heading 3a") > 0 And InStr(pg.Previous.Previous.Style, "Heading 2") > 0) _
                    Or (pg.Previous.Range.Font.Size = Style.Font.Size And pg.Previous.Range.Font.Bold = Style.Font.Bold) Then
                        rgPg.ListFormat.ApplyListTemplate rgPg.ListFormat.ListTemplate, False, wdlistthispointforward
                        rgPg.ListFormat.ApplyListTemplate ListTemplate:=rgPg.ListFormat.ListTemplate, ContinuePreviousList:=False, applyto:=wdlistthispointforward
                    End If
                End If
            Next pg
        End If
        On Error GoTo 0
        
        'parse date
        Dim dSigning As Date    '->Document Date
        Dim dCommence As Date   '->Publishing End Date
        Dim dBelowIssuingAuth As Date   '->Effective Date
        Dim emptyDate As Date
        dSigning = emptyDate
        dCommence = emptyDate
        dBelowIssuingAuth = emptyDate
        'Date of Signing, from cc
         If docOri.SelectContentControlsByTitle("DateOfSigning").Count > 0 Then
            Set ccOri = docOri.SelectContentControlsByTitle("DateOfSigning").Item(1)
            If IsDate(ccOri.Range.Text) Then
                dSigning = CDate(ccOri.Range.Text)
            End If
        End If
        'Date in cc Commencement, from keyword "force on"
        If docOri.SelectContentControlsByTitle("Commencement").Count > 0 Then
            Set ccOri = docOri.SelectContentControlsByTitle("Commencement").Item(1)
            Dim dStr As String
            dStr = Replace(ccOri.Range.Paragraphs(2).Range.Text, vbCr, "")
            dStr = Replace(dStr, ".", "")
            dStr = Application.CleanString(dStr)
            If InStr(dStr, "force on") > 0 Then
                dStr = Right(dStr, Len(dStr) - InStr(dStr, "force on") - 7)
                If IsDate(dStr) Then
                    dCommence = CDate(dStr)
                End If
            End If
        End If
        'Date below Issuing Auth, from keyword "Dated at Wellington this", between CC IssuingPerson and CC IssuingAuthority
        If docOri.SelectContentControlsByTitle("IssuingAuthority").Count > 0 And docOri.SelectContentControlsByTitle("IssuingPerson").Count > 0 _
        And docTpl.SelectContentControlsByTitle("IssuingAuthority").Count > 0 And docTpl.SelectContentControlsByTitle("Approver").Count > 0 Then
            Dim ccIA As ContentControl
            Dim ccIP As ContentControl
            Set ccIA = docOri.SelectContentControlsByTitle("IssuingAuthority").Item(1)
            Set ccIP = docOri.SelectContentControlsByTitle("IssuingPerson").Item(1)
            If ccIA.Range.End < ccIP.Range.Start Then
                Set rgOri = ccIA.Range
                rgOri.SetRange ccIA.Range.End, ccIP.Range.Start
                With rgOri.Find
                    .ClearFormatting
                    .ClearAllFuzzyOptions
                    .MatchCase = False
                    .Wrap = wdFindStop
                    .MatchWholeWord = True
                    .Text = "Dated at Wellington this"
                    .Execute
                End With
                Set rgTpl = docTpl.Range
                Dim ccIATpl As ContentControl
                Dim ccApprover As ContentControl
                Set ccIATpl = docTpl.SelectContentControlsByTitle("IssuingAuthority").Item(1)
                Set ccApprover = docTpl.SelectContentControlsByTitle("Approver").Item(1)
                With rgTpl.Find
                    .ClearFormatting
                    .ClearAllFuzzyOptions
                    .MatchCase = False
                    .Wrap = wdFindStop
                    .MatchWholeWord = True
                    .Text = "Dated at Wellington this"
                    .Execute
                End With
                If rgOri.Find.Found And rgOri.Start > ccIA.Range.End And rgOri.End < ccIP.Range.Start _
                And rgTpl.Find.Found And rgTpl.Start > ccIATpl.Range.End And rgTpl.End < ccApprover.Range.Start Then
                    rgOri.Expand wdParagraph
                    rgTpl.Expand wdParagraph
                    rgTpl.Text = Replace(rgOri.Text, "  ", " ")
                    Log "Set Date of signing in body"
                End If
                
                Set rgTpl = Nothing
                Set rgOri = Nothing
            End If
        End If
        'set value in DIP
        For Each mPty In docTpl.ContentTypeProperties
            If mPty.Name = "Document Date" Then
                If dSigning <> emptyDate Then
                    mPty.Value = dSigning
                    Log "Set Document Date" & vbTab & dSigning
                End If
            End If
            If mPty.Name = "Date of Signing" Then
                If dSigning <> emptyDate Then
                    mPty.Value = dSigning
                    Log "Set Date of Signing" & vbTab & dSigning
                End If
            End If
            If mPty.Name = "Effective Date" Then
                If dCommence <> emptyDate Then
                    mPty.Value = dCommence
                    Log "Set Effective Date" & vbTab & dBelowIssuingAuth
                End If
            End If
        Next mPty
        
        'delete restricted edit area, tao@allfields.co.nz 5/12/2015
        On Error Resume Next
        For Each pr In docOri.Paragraphs
            If pr.Range.Editors.Count > 0 Then
                For iEditor = 1 To pr.Range.Editors.Count
                    pr.Range.Editors(iEditor).Delete
                Next iEditor
            End If
        Next pr
                
        'update TOC
        If docTpl.TablesOfContents.Count > 0 Then
            docTpl.TablesOfContents(1).Update
        End If

        Dim sFileName As String
        sFileName = Right(docTpl.Name, Len(docTpl.Name) - InStrRev(docTpl.Name, "."))
        sFileName = doc.Path & "\Processed\" & docOri.Name
        
        docTpl.SaveAs2 FileName:=sFileName, fileformat:=wdFormatXMLDocument, CompatibilityMode:=15
        iFileExported = iFileExported + 1
        Log "Exported: " & vbTab & doc.Path & "\Processed\" & docTpl.Name & vbNewLine
        docTpl.Close SaveChanges:=False
        docOri.Close SaveChanges:=False
        Set docOri = Nothing
        Set docTpl = Nothing
        errMsg = ""
OpenFail:
        If errMsg <> "" Then
            Log ("!!!!!Open failed: " & vbTab & errMsg & vbNewLine)
            If Not docTpl Is Nothing Then docTpl.Close False
            If Not docOri Is Nothing Then docOri.Close False
        End If
        
        sFile = Dir() 'next file
        If sFile = ".." Then sFile = ""
        DoEvents
    Loop
    
    Log iFileOpened & " file(s) opened and " & vbTab & iFileExported & " file(s) exported" & vbNewLine
    Set docTpl = Nothing
    Set docOri = Nothing
    MsgBox iFileOpened & " files opened and " & vbTab & iFileExported & " files exported"
    Debug.Print iFileOpened & " files opened and " & vbTab & iFileExported & " files exported"
    'clear clip board
    Dim MyData As DataObject
    Set MyData = New DataObject
    MyData.SetText ""
    MyData.PutInClipboard
NoCC:
    Log ("!!!!!No CC: " & vbTab & errMsg & vbTab & Err.Description)
    If Not docTpl Is Nothing Then docTpl.Close False
    If Not docOri Is Nothing Then docOri.Close False
    Exit Sub
End Sub



Sub tmp()
    Dim pty As DocumentProperty
    Dim mPty As MetaProperty
    Dim doc As Document
    Set doc = ThisDocument
    
    For Each pty In doc.BuiltInDocumentProperties
        Debug.Print pty.Name
    Next pty
    
    For Each pty In doc.CustomDocumentProperties
        Debug.Print pty.Name
    Next pty
    
    For Each mPty In doc.ContentTypeProperties
        Debug.Print mPty.Name
    Next mPty
End Sub

Public Function NumberEnding(i As Integer) As String
    NumberEnding = Switch(i Mod 100 >= 11 And i Mod 100 <= 13, "th", _
                i Mod 10 = 1, "st", _
                i Mod 10 = 2, "nd", _
                i Mod 10 = 3, "rd", _
                1 = 1, "th")
End Function

Sub testnet()
    Debug.Print ThisDocument.BuiltInDocumentProperties.Count

    Const sTplFile As String = "https://cohesiondemo-piritahi.intergen.net.nz/sites/rgr/WordTemplate/"
    
    'check if logged in to SP
    Dim oHttpRequest
    Err.Clear
    Set oHttpRequest = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    With oHttpRequest
        .Open "GET", sTplFile, False
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Pragma", "no-cache"
        .send
    End With
    Debug.Print Err.Number & Err.Description
End Sub

Sub Util()
    Dim docTpl As Document
    'sort out numbering
    Dim pg As Paragraph
    Dim rgPg As Range
    Dim Style As Style
    Set docTpl = Documents("IHS RETURNAP.ALL.docx")
    Set Style = docTpl.Styles("Heading 2a")
    
    On Error Resume Next
    For Each pg In docTpl.Paragraphs
        Set rgPg = pg.Range
        If pg.Style = "Clause L1" And rgPg.ListParagraphs.Count > 0 Then
            If pg.Previous.Range.Style = "Heading 2a" _
                Or pg.Previous.Style = "Heading 2" _
                Or (pg.Previous.Range.Font.Size = Style.Font.Size And pg.Previous.Range.Font.Bold = Style.Font.Bold) Then
                rgPg.ListFormat.ApplyListTemplate rgPg.ListFormat.ListTemplate, False, wdlistthispointforward
            End If
        End If
    Next pg
End Sub
