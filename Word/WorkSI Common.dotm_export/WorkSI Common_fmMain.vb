VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmMain 
   Caption         =   "Work SI Template"
   ClientHeight    =   10410
   ClientLeft      =   165
   ClientTop       =   585
   ClientWidth     =   13185
   OleObjectBlob   =   "WorkSI Common_fmMain.frx":0000
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' These templates have been prepared and developed for WorkSafe Investigation
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             May 2017
' Description:      Form used for picking template to load.
'-----------------------------------------------------------------------------

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOK_Click()
    Dim tmrStart As Long
    tmrStart = Timer
    Dim tmrTmp As Long
    'validation
    If Not CheckTB(Me.fmInfo) Then
        MsgBox "Please enter content for this field."
        Exit Sub
    End If
    
    If imgLogo.Tag = "" Then
        If MsgBox("Would you like to load a logo image?", vbYesNo) = vbYes Then
            imgLogo_MouseUp 1, 1, 1, 1
        End If
    End If
        
    'check if any template has been selected
    Dim HasTemplate As Boolean
    For i = 0 To UBound(Blocks)
        If Blocks(i).Selected Then
            HasTemplate = True
            Exit For
        End If
    Next i
    If Not HasTemplate Then
        MsgBox "Please select at leat one template."
        Exit Sub
    End If
    
    'confirm
    If MsgBox("Generate the document now based on your selections?", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    Me.Hide
    Dim blHide As Boolean   'flag whether hide main document during processing
    blHide = False
    Dim doc As Document
    Set doc = ActiveDocument
    If blHide Then
        Dim docTmp As Document  'dummy document to show when hiding the main one
        Set docTmp = Documents.Add
        docTmp.Activate
    End If
    'to reduce processing time, make it invisible first
    doc.Windows(1).Visible = Not blHide
    Application.ScreenUpdating = False
    
    'write back custom values
    WriteCP Me.txtAddress
    WriteCP Me.txtCompany
    WriteCP Me.txtContractor
    WriteCP Me.txtDate
    WriteCP Me.txtEmail
    WriteCP Me.txtOfficer
    WriteCP Me.txtPhone

    'insert building blocks
    If Len(doc.Content.Text) > 1 Then
        If MsgBox("Would you like to delete content in document and create a new one?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
      
    Me.Hide
    doc.Content.Delete  'cleare content
    Dim rg As Range
    Dim rgStart As Long 'start position before next BB insertion
    Dim rgTmp As Range
    Dim cntParsed As Integer
    Dim blFirst As Boolean
    tmrTmp = Timer      'set timer to measure each BB insertion
    For i = 0 To UBound(Blocks)
        Set rg = doc.Content
        rgStart = rg.End
        rg.Collapse wdCollapseEnd     'set insert point to start
        If Blocks(i).Selected Then
            doc.AttachedTemplate.BuildingBlockEntries(Blocks(i).Name).Insert rg, True
            'fix numbering issue: restart numbering of 1st paragraph then continue numbering for the rest
            rg.SetRange rgStart, doc.Content.End
            'work around to correct layout formatting and/or numbering in sections
            Dim tm5E
            tm5E = Timer
            On Error Resume Next
            Select Case LCase(Left(Blocks(i).Name, 2))
                Case "1a", "1b"
                    If rg.Tables.Count > 0 Then
                        ParaIndentment rg.Tables(1).Range
                    End If
                Case "4f"
                    If rg.Tables.Count > 0 Then
                        rg.Tables(1).Range.Style = "Normal"
                        rg.Tables(1).Cell(1, 1).Range.Font.Bold = True
                    End If
                Case "5b"
                    If rg.Tables.Count > 3 Then
                        ParaIndentment rg.Tables(1).Range
                        ParaIndentment rg.Tables(3).Range
                    End If
                Case "5c"
                    If rg.Tables.Count > 4 Then
                        ParaIndentment rg.Tables(1).Range
                        rg.Tables(1).Cell(1, 1).Range.Font.Bold = True
                        rg.Tables(1).Cell(2, 1).Range.Font.Bold = True
                        ParaIndentment rg.Tables(3).Range
                        ParaIndentment rg.Tables(5).Range
                    End If
                Case "5e"
                    If rg.Tables.Count > 0 Then
'                        'table 1
'                        rg.Tables(1).Range.Style = "Normal"
'                        rg.Tables(1).Range.ParagraphFormat.SpaceAfter = 0
'                        rg.Tables(1).Cell(1, 1).Range.Style = "Note"
'                        'table 3
'                        rg.Tables(3).Range.Style = "Normal"
'                        rg.Tables(3).Rows(1).Range.ParagraphFormat.SpaceAfter = 0
'                        rg.Tables(3).Rows(2).Range.ParagraphFormat.SpaceAfter = 0
'                        rg.Tables(3).Rows(3).Range.ParagraphFormat.SpaceAfter = 0
'                        rg.Tables(3).Rows(4).Range.ParagraphFormat.SpaceAfter = 0
'                        rg.Tables(3).Rows(3).Range.Font.Bold = True
'                        rg.Tables(3).Rows(6).Range.Font.Bold = True
                        'sort out numbering
                        For cntTb = 1 To rg.Tables.Count
                            tm5Etb = Timer
                            DoEvents
                            Dim rgTb As Range
                            Set rgTb = rg.Tables(cntTb).Range
                            Dim cl As Cell
                            For Each cl In rgTb.Cells
                                If cl.Range.ListParagraphs.Count > 0 Then
                                    For cPr = 1 To cl.Range.ListParagraphs.Count
                                        DoEvents
                                        Set rgTmp = cl.Range.ListParagraphs(cPr).Range
                                        rgTmp.Collapse wdCollapseStart
                                        rgTmp.Style = NUMBER_LIST_2_STYLE 'rgTmp.Style   'set style to correct potential issue
                                        If cPr = 1 Then
                                            If Not rgTmp.ListFormat.ListTemplate Is Nothing Then
                                                rgTmp.ListFormat.ApplyListTemplate rgTmp.ListFormat.ListTemplate, False
                                            End If
                                            rgTmp.ParagraphFormat.Alignment = wdAlignParagraphLeft
                                        End If
                                    Next cPr
                                End If
                            Next cl
                            Debug.Print "5e tables: " & Timer - tm5Etb
                        Next cntTb
                    End If
                    Debug.Print "5e: " & Timer - tm5E
                Case "5f"
                    If rg.Tables.Count > 0 Then
                        ParaIndentment rg.Tables(1).Range
                    End If
                Case "5a", "7b"
                    tm7 = Timer
                    If rg.ListParagraphs.Count > 0 Then
                        cntParsed = 0
                        blFirst = True
                        Debug.Print Blocks(i).Name
                        For j = 1 To rg.Paragraphs.Count
                            DoEvents
                            If rg.Paragraphs(j).Range.ListParagraphs.Count > 0 Then
                                'If rg.Paragraphs(j).Range.ListFormat.ListType = wdListOutlineNumbering And IsNumeric(rg.Paragraphs(j).Range.ListFormat.ListString) Then
                                If IsNumeric(rg.Paragraphs(j).Range.ListFormat.ListString) Then
                                    Set rgTmp = rg.Paragraphs(j).Range
                                    rgTmp.Collapse
                                    rgTmp.Style = rgTmp.Style 'NUMBER_LIST_STYLE     '#TBT: apply style can set numbering correct
                                    If blFirst Then
                                        If Not rgTmp.ListFormat.ListTemplate Is Nothing Then
                                            rgTmp.ListFormat.ApplyListTemplate rgTmp.ListFormat.ListTemplate, ContinuePreviousList:=False 'IIf(blFirst, False, True)
                                        End If
                                        blFirst = False
                                    End If
                                End If
                            End If
                        Next j
                        Debug.Print "7: " & Timer - tm7
                    End If
                Case "7c"
                    If rg.Tables.Count > 5 Then
                        ParaIndentment rg.Tables(1).Range
                        rg.Tables(1).Range.Font.Bold = True
                        ParaIndentment rg.Tables(3).Range
                        rg.Tables(3).Range.Font.Bold = True
                        ParaIndentment rg.Tables(4).Range
                        ParaIndentment rg.Tables(6).Range
                    End If
                Case "7d"
                    If rg.Tables.Count > 3 Then
                        ParaIndentment rg.Tables(1).Range
                        rg.Tables(1).Range.Font.Bold = True
                        ParaIndentment rg.Tables(3).Range
                        ParaIndentment rg.Tables(4).Range
                    End If
                Case "8b"
                    If rg.Tables.Count > 3 Then
                        ParaIndentment rg.Tables(1).Range
                    End If
            End Select
            'End If
        End If
    Next i
    Debug.Print "Insert: " & Timer - tmrTmp
    'find two consective section breaks and make it one
    'there's a paragraph mark in between
    tmrTmp = Timer  'timer to measure find & replace section breaks

    Set rg = doc.Content
    rg.Collapse wdCollapseStart
    With rg.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .ClearHitHighlight
        .Forward = True
        .Wrap = wdFindContinue
        .Text = "^b^p^p^b"    'section break/paragraph mark/section break
        .Execute
        Do While .Found
            rg.SetRange rg.Start + 1, rg.End 'rg.Start + 1, rg.End   'VBA not allows replace with '^b' char, so instead we delete a section break char
            'rg.Text = ""
            rg.Delete
            .Execute
        Loop
    End With
    Debug.Print "Delete section breaks: " & Timer - tmrTmp
    
    'delete section break on start of 1st page
    If doc.Paragraphs(1).Range.Text = Chr(12) Then 'chr(12) is section break char
        doc.Paragraphs(1).Range.Delete
    End If
    
    '###TODO: insert logo image
    Dim hf As HeaderFooter
    Dim SCT As Section
    tmrTmp = Timer  'timer to measure insert logo
    For i = 1 To doc.Sections.Count
        Set SCT = doc.Sections(i)
        If SCT.Headers(wdHeaderFooterEvenPages).Exists Then
            ReplacePicInHeader SCT.Headers(wdHeaderFooterEvenPages)
        End If
        If SCT.Headers(wdHeaderFooterFirstPage).Exists Then
            ReplacePicInHeader SCT.Headers(wdHeaderFooterFirstPage)
        End If
        If SCT.Headers(wdHeaderFooterPrimary).Exists Then
            ReplacePicInHeader SCT.Headers(wdHeaderFooterPrimary)
        End If
    Next i
    Debug.Print "Insert logo: " & Timer - tmrTmp
    
    'update fields
    doc.Fields.Update
    Set rg = doc.Content
    rg.Collapse wdCollapseStart
    rg.Select
    If blHide Then
        docTmp.Close False
    End If
    doc.Windows(1).Visible = True
    Application.ScreenUpdating = True
    doc.Activate
    Debug.Print Timer - tmrStart
End Sub

Private Sub imgLogo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim img As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .ButtonName = "OK"
        .Title = "Select an image for logo"
        .Filters.Clear
        .Filters.Add "JPEG", "*.JPG"
        .Filters.Add "JPEG File Interchange Format", "*.JPEG"
        .Filters.Add "Bitmap", "*.BMP"
        .Filters.Add "Graphics Interchange Format", "*.GIF"
        .Filters.Add "Portable Network Graphics", "*.PNG"
        .Filters.Add "Tag Image File Format", "*.TIFF"

        If .Show = -1 Then
            img = .SelectedItems(1)
            'imgLogo.AutoSize = False
            'imgLogo.PictureSizeMode = fmPictureSizeModeZoom
            'imgLogo.PictureAlignment = fmPictureAlignmentCenter
        End If
    End With
    Set Me.imgLogo.Picture = LoadPicture(img) '("C:\Users\tao\Pictures\zoolander-for-ants-what-is-this-a-museum-for-ants.jpg", imgLogo.Width, imgLogo.Height)
    imgLogo.Tag = img
    
    'refresh image control
    Me.Repaint
End Sub


Private Sub UserForm_Initialize()
    Dim SectionName() As String
    SectionName = Split(SECTION_NAMES, ",")
    Dim v As Variant
    Me.fmSection.Controls.Clear
    For i = 1 To UBound(SectionName) + 1
        Dim cb As MSForms.CheckBox
        Set cb = fmSection.Controls.Add("Forms.Checkbox.1", "cb1")
        With cb
            .Top = IIf(i Mod 2 = 0, TOP_GAP * (i / 2), TOP_GAP * ((i + 1) / 2)) - 12
            .Left = IIf(i Mod 2 = 0, LEFT_COLUMN_2, LEFT_COLUMN_1)
            .Width = WIDTH_SECTION
            .Height = HEIGHT_SECTION
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE
            .Caption = SectionName(i - 1)
            .Tag = Left(.Caption, 2)
        End With
        ReDim Preserve cbSections(i - 1)
        Set cbSections(i - 1).cb = cb
        cbSections(i - 1).Caption = cb.Caption
    Next i
    
    'populate building blocks array for later use
    Dim doc As Document
    Set doc = ThisDocument
    Dim tmp As Template
    Set tmp = doc.AttachedTemplate
    If tmp.BuildingBlockEntries.Count = 0 Then
        MsgBox "No building blocks found."
        End
    End If
    
    For i = 1 To tmp.BuildingBlockEntries.Count
        ReDim Preserve Blocks(i - 1)
        Blocks(i - 1).Name = tmp.BuildingBlockEntries(i).Name
        Blocks(i - 1).Selected = False
        Blocks(i - 1).Description = tmp.BuildingBlockEntries(i).Description
    Next i
    
    'sort on name
    'put '1a' before '10'
    Dim bb As Block
    Dim iNum As String  'digit of item i
    Dim jNum As String  'digit of item j
    For i = 0 To UBound(Blocks)
        For j = i + 1 To UBound(Blocks)
            iNum = ParseFileName(Blocks(i).Name)
            jNum = ParseFileName(Blocks(j).Name)
            '####If Blocks(j).Name < Blocks(i).Name Then
            If StrComp(jNum, iNum) < 0 Then
                Set bb = Blocks(j)
                Set Blocks(j) = Blocks(i)
                Set Blocks(i) = bb
            End If
        Next j
    Next i
    
    'hide scroll bar
    Me.fmTemplates.ScrollBars = fmScrollBarsNone
    Me.fmSelected.ScrollBars = fmScrollBarsNone
    
    'read in custom properties
    On Error Resume Next
    txtCompany.Text = ReadCP(txtCompany)
    txtContractor.Text = ReadCP(txtContractor)
    txtOfficer.Text = ReadCP(txtOfficer)
    txtAddress.Text = ReadCP(txtAddress)
    txtEmail.Text = ReadCP(txtEmail)
    txtPhone.Text = ReadCP(txtPhone)
    txtDate.Text = Format(Date, DATE_FORMAT)
End Sub
