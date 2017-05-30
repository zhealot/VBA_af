Attribute VB_Name = "RibbonCode"
Public OffsetPages As String

Sub insertLandscapePage(control As IRibbonControl)
    Dim doc As Document
    Dim RgTmp As Range
    Dim blLastPage As Boolean
    Dim ftr As HeaderFooter
    Dim iCurrentSection As Integer
    Dim blLandscapePage As Boolean
    Dim strBuildingBlockItem As String
    Dim iPageNum As Integer
    Dim LinkToPrevious As Boolean
    Dim RestartingNum As Boolean
            
    Application.ScreenUpdating = False
    LinkToPrevious = False
    RestartingNum = False
    Set doc = ActiveDocument
    blLastPage = IIf(Selection.Information(wdActiveEndPageNumber) = Selection.Information(wdNumberOfPagesInDocument), True, False)
    iCurrentSection = Selection.Information(wdActiveEndSectionNumber)
    iPageNum = Selection.Information(wdActiveEndPageNumber)
    With Selection
        If .PageSetup.Orientation = wdOrientLandscape Then 'cursor on a landscape page
            .Bookmarks("\page").Range.Select
            .Collapse wdCollapseEnd
            .MoveLeft wdCharacter, 1
            .InsertBreak wdPageBreak
            If .Information(wdActiveEndPageNumber) = iPageNum Then 'if page starts with a section break, insert 2 page break to generate a new page, else needs only one
                .InsertBreak Type:=wdPageBreak
            End If
            LinkToPrevious = False
        Else    'cursor on a portrait page
            iCurrentSection = iCurrentSection + 1
            .Bookmarks("\page").Range.Select
            .Collapse wdCollapseEnd
            iPageNum = Selection.Information(wdActiveEndPageNumber)
            If .PageSetup.Orientation = wdOrientLandscape Then
                .InsertBreak wdPageBreak
                If .Information(wdActiveEndPageNumber) = iPageNum Then 'if page starts with a section break, insert 2 page break to generate a new page, else needs only one
                    .InsertBreak Type:=wdPageBreak
                End If
            Else
                .InsertBreak Type:=wdSectionBreakNextPage
                If Not blLastPage Then
                    .InsertBreak Type:=wdSectionBreakNextPage
                End If
            End If
        End If
        If Not blLastPage Then
            .GoToPrevious wdGoToPage
        End If
    End With
    
    With Selection.PageSetup
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(3)
        .BottomMargin = CentimetersToPoints(3)
        .LeftMargin = CentimetersToPoints(1.5)
        .RightMargin = CentimetersToPoints(3)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.27)
        .FooterDistance = CentimetersToPoints(1.3)
        .OddAndEvenPagesHeaderFooter = True
        .MirrorMargins = True
    End With
    
    If iCurrentSection < doc.Sections.Count Then
        If doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterFirstPage).Exists Then
            doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterFirstPage).LinkToPrevious = False
            doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterFirstPage).PageNumbers.RestartNumberingAtSection = False
        End If
        If doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterPrimary).Exists Then
            doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
            doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = False
        End If
        If doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterEvenPages).Exists Then
            doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterEvenPages).LinkToPrevious = False
            doc.Sections(iCurrentSection + 1).Footers(wdHeaderFooterEvenPages).PageNumbers.RestartNumberingAtSection = False
        End If
    End If
    
    If doc.Sections(iCurrentSection).Footers(wdHeaderFooterPrimary).Exists Then
        Set ftr = doc.Sections(iCurrentSection).Footers(wdHeaderFooterPrimary)
        Call fFixFooter(ftr, True, LinkToPrevious, RestartingNum)
    End If
    
    If doc.Sections(iCurrentSection).Footers(wdHeaderFooterEvenPages).Exists Then
        Set ftr = doc.Sections(iCurrentSection).Footers(wdHeaderFooterEvenPages)
        Call fFixFooter(ftr, False, LinkToPrevious, RestartingNum)
    End If
    Application.ScreenUpdating = True
End Sub

Sub FixFooters(control As IRibbonControl)
    Dim sSection As Section
    Dim dDoc As Document
    Dim blRestartNumbering As Boolean
    Dim blLinkToPrevious As Boolean
    Dim ftFooter As HeaderFooter
    
    Application.ScreenUpdating = False
    Set dDoc = ActiveDocument
    
    'assign value for offset pages
    If dDoc.Sections.Count > 1 Then
        Dim rgSct2 As Range
        Set rgSct2 = dDoc.Sections(2).Range
        rgSct2.Collapse wdCollapseStart
        OffsetPages = InputBox(rgSct2.Information(wdActiveEndPageNumber) - 1 & " pages in Section 1" & _
            vbNewLine & "Please enter new value", "Offset", dDoc.BuiltInDocumentProperties("Comments"))
        If OffsetPages = "" Or Not IsNumeric(OffsetPages) Then
            OffsetPages = rgSct2.Information(wdActiveEndPageNumber) - 1
        End If
    Else
        OffsetPages = 0
    End If
    dDoc.BuiltInDocumentProperties("Comments") = OffsetPages
    If dDoc.TrackRevisions Then
        If MsgBox("Tracking Changes is on, continue?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    For Each sSection In dDoc.Sections
        With sSection
            If sSection.PageSetup.Orientation = wdOrientPortrait Then
                .PageSetup.TopMargin = CentimetersToPoints(3)
                .PageSetup.BottomMargin = CentimetersToPoints(3)
                .PageSetup.LeftMargin = CentimetersToPoints(3)
                .PageSetup.RightMargin = CentimetersToPoints(1.5)
                .PageSetup.MirrorMargins = True
            End If
            If sSection.PageSetup.Orientation = wdOrientLandscape Then
                .PageSetup.TopMargin = CentimetersToPoints(3)
                .PageSetup.BottomMargin = CentimetersToPoints(3)
                .PageSetup.LeftMargin = CentimetersToPoints(1.5)
                .PageSetup.RightMargin = CentimetersToPoints(3)
                .PageSetup.MirrorMargins = True
            End If
        End With
    
        If sSection.Footers(wdHeaderFooterPrimary).Exists Then
            Set ftFooter = sSection.Footers(wdHeaderFooterPrimary)
            If ftFooter.Range.Fields.Count > 0 And InStr(ftFooter.Range.Text, "PAGE") > 0 Then
                blRestartNumbering = ftFooter.PageNumbers.RestartNumberingAtSection
                Call fFixFooter(ftFooter, True, False, blRestartNumbering)
            End If
        End If
        If sSection.Footers(wdHeaderFooterEvenPages).Exists Then
            Set ftFooter = sSection.Footers(wdHeaderFooterEvenPages)
            If ftFooter.Range.Fields.Count > 0 And InStr(ftFooter.Range.Text, "PAGE") > 0 Then
                blRestartNumbering = ftFooter.PageNumbers.RestartNumberingAtSection
                Call fFixFooter(ftFooter, False, False, blRestartNumbering)
            End If
        End If
    Next sSection
    MsgBox "Done"
    Application.ScreenUpdating = True
End Sub

'Callback for customButton2 onAction
Sub FixTalbes(control As IRibbonControl)
    Dim tTb As Table
    Dim dDoc As Document
    Dim cCl As Cell
    
    Application.ScreenUpdating = False
    Set dDoc = ActiveDocument
    
    For Each tTb In dDoc.Tables
        tTb.AutoFitBehavior wdAutoFitWindow
        For Each cCl In tTb.Range.Cells
            If cCl.RowIndex = 1 Then
                With cCl.Borders(wdBorderBottom)
                    .Visible = True
                    .LineStyle = wdLineStyleNone
                    .Color = Options.DefaultBorderColor
                    If .LineWidth > wdLineWidth100pt Then
                        .LineWidth = wdLineWidth100pt
                    End If
                End With
            End If
        Next cCl
    Next tTb
    MsgBox "Done"
    Application.ScreenUpdating = True
End Sub

'Callback for customButton3 onAction
Sub FixFields(control As IRibbonControl)
    Dim TOC As TableOfContents
    Dim dDoc As Document
    Dim fFld As Field
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    Set dDoc = ActiveDocument
    
    'update table of content
    If dDoc.TablesOfContents.Count > 0 Then
        For i = 1 To dDoc.TablesOfContents.Count Step 1
            dDoc.TablesOfContents(i).Update
        Next i
    End If
    
    'update table of figure
    If dDoc.TablesOfFigures.Count > 0 Then
        For i = 1 To dDoc.TablesOfFigures.Count Step 1
            dDoc.TablesOfFigures(i).Update
        Next i
    End If
        
    Dim RgStr As Range
    For Each RgStr In dDoc.StoryRanges
        RgStr.Fields.Update
    Next RgStr
    
    Dim sSct As Section
    
    For Each sSct In dDoc.Sections
        If sSct.Footers(wdHeaderFooterEvenPages).Exists Then
            sSct.Footers(wdHeaderFooterEvenPages).Range.Fields.Update
        End If
        If sSct.Footers(wdHeaderFooterFirstPage).Exists Then
            sSct.Footers(wdHeaderFooterFirstPage).Range.Fields.Update
        End If
        If sSct.Footers(wdHeaderFooterPrimary).Exists Then
            sSct.Footers(wdHeaderFooterPrimary).Range.Fields.Update
        End If
    Next sSct
    MsgBox "Done"
    Application.ScreenUpdating = True
End Sub


Sub ReplaceTerms(control As IRibbonControl)
    Dim dDoc As Document
    Dim RgTmp As Range
    Dim tb As Table
    Dim rw As Row
    Dim sFrom As String
    Dim sTo As String
    Dim sMsg As String
    Dim sTmp As String
    Dim i As Integer
    Dim tmStart As Single
            
    Application.ScreenUpdating = True
    Set dDoc = ActiveDocument
    
    If ThisDocument.Tables.Count > 0 Then
        If InStr(ThisDocument.Tables(IIf(control.Tag = "proof", 1, 2)).Rows(1).Cells(1).Range.Text, IIf(control.Tag = "proof", "Proof for terms button", "Check content needed button")) > 0 Then
            Set tb = ThisDocument.Tables(IIf(control.Tag = "proof", 1, 2))
            For Each rw In tb.Rows
                If rw.Index > 1 Then
                    If Not Trim(rw.Cells(1).Range.Text) = "" And Not Trim(rw.Cells(2).Range.Text) = "" Then
                        i = 0
                        sFrom = Trim(Replace(Replace(Application.CleanString(rw.Cells(1).Range.Text), Chr(13), ""), Chr(9), ""))
                        sTo = Trim(Replace(Replace(Application.CleanString(rw.Cells(2).Range.Text), Chr(13), ""), Chr(9), ""))
                        Set RgTmp = dDoc.Range
                        tmStart = Timer
                        While ReplaceString(RgTmp, sFrom, sTo)
                            If Timer - tmStart > 30 Then
                                If MsgBox("30 seconds passed, continue?", vbYesNo) = vbNo Then
                                    Exit Sub
                                End If
                            End If
                            i = i + 1
                            RgTmp.SetRange RgTmp.End, dDoc.Range.End
                        Wend
                        If i > 0 Then
                            sMsg = sMsg & vbNewLine & i & vbTab & sFrom & vbTab & " ->" & vbTab & sTo
                        End If
                    End If
                End If
            Next rw
            MsgBox sMsg
        Else
            MsgBox "No proof for tems table found in template"
        End If
    Else
        MsgBox "No table found in template"
    End If
End Sub


'Callback for customButton4 onAction
Sub FinalCheck(control As IRibbonControl)
    Dim dDoc As Document
    Dim RgTmp As Range
    
    Application.ScreenUpdating = False
    Set dDoc = ActiveDocument
    Set RgTmp = dDoc.Range
    With RgTmp.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .MatchWholeWord = False
        .MatchCase = True
        .Forward = True
        .Wrap = wdFindStop
        .Text = "Error! Reference source not found"
        .Execute
        If .Found Then
            RgTmp.Select
            MsgBox "Invalid Reference"
            Exit Sub
        End If
    End With
    'check empty title and subject
    If Trim(dDoc.BuiltInDocumentProperties("Subject")) = "VOL XXX SECTION [XXX]" Then
        MsgBox "Document Subject is: VOL XXX SECTION [XXX]"
    End If
    
    
    
    'check if last page is odd page
    If dDoc.BuiltInDocumentProperties("Number of pages") Mod 2 = 1 Then
        Dim RgLst As Range
        Set RgLst = dDoc.Paragraphs.Last.Range
        RgLst.Collapse wdCollapseEnd
        RgLst.InsertBreak
        RgLst.InsertAfter "This page is left intentionally blank"
        MsgBox "One page is inserted to the end of this document"
    End If
    MsgBox "Done"
    Application.ScreenUpdating = True
End Sub


'function to fix a footer
Function fFixFooter(ft As HeaderFooter, IsPrimary As Boolean, LinkToPrevious As Boolean, RestartNumbering As Boolean)
    Dim blReNum As Boolean
    Dim blLkPr As Boolean
    Dim RgTmp As Range
    Dim sFooterName As String
    
    ft.PageNumbers.RestartNumberingAtSection = RestartNumbering
    ft.Range.Delete
    If ft.Parent.PageSetup.Orientation = wdOrientLandscape Then
        If ft.Parent.Index > 1 Then
            If ActiveDocument.Sections(ft.Parent.Index - 1).PageSetup.Orientation = wdOrientPortrait Then
                LinkToPrevious = False
            End If
        End If
        If IsPrimary Then
            sFooterName = "NX2 landscape footer odd"
        Else
            sFooterName = "NX2 landscape footer even"
        End If
    End If
    If ft.Parent.PageSetup.Orientation = wdOrientPortrait Then
        If ft.Parent.Index > 1 Then
            If ActiveDocument.Sections(ft.Parent.Index - 1).PageSetup.Orientation = wdOrientLandscape Then
                LinkToPrevious = False
            End If
        End If
        If IsPrimary Then
            sFooterName = "NX2 portrait footer odd"
        Else
            sFooterName = "NX2 portrait footer even"
        End If
    End If
    Err.Clear
    On Error Resume Next
    Application.Templates(ThisDocument.Path & "\" & ThisDocument.Name).BuildingBlockEntries(sFooterName).Insert where:=ft.Range, RichText:=True
    If Err.Number > 0 Then
    '### keep log
    End If
    If Not ft.Range Is Nothing Then
'        If Replace(Application.CleanString(ft.Range.Paragraphs.Last.Range.Text), Chr(13), "") = "" Then
'            Set RgTmp = ft.Range
'            RgTmp.SetRange ft.Range.Paragraphs(ft.Range.Paragraphs.Count - 1).Range.End, RgTmp.End
'            RgTmp.Delete
'        End If
        ft.Range.Paragraphs.Last.Range.Font.Size = 1
    End If
    ft.LinkToPrevious = LinkToPrevious
    ft.PageNumbers.RestartNumberingAtSection = RestartNumbering
    ft.Range.Fields.Update
End Function

Public Function NumInString(str As String) As String
    NumInString = ""
    str = Application.CleanString(str)
    If str = "" Then
        Exit Function
    End If
    Dim i As Integer
    For i = Len(str) To 1 Step -1
        If IsNumeric(Mid(str, i, 1)) Then
            NumInString = Mid(str, i, 1) & NumInString
        End If
    Next i
End Function

Sub test()
    Dim rg As Range
    
    Set rg = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range '.Fields(3).Code
    Debug.Print rg.Text
    With rg.Find
        .ClearFormatting
        .MatchCase = False
        .MatchWholeWord = False
        .ClearAllFuzzyOptions
        .Format = False
        .Text = "FOR" 'NumInString(rg.Text)
        .Execute
    End With
    If rg.Find.Found Then
        rg.Find.Replacement.Text = "314"
        rg.Find.Execute
    End If
    ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Fields(3).Update
End Sub

Function ReplaceString(rg As Range, sFrom As String, sTo As String) As Boolean
    ReplaceString = False
    'Dim rgT As Range
    Dim lEnd As Long
    Dim lTmpStart As Long
    Dim lTmpEnd As Long
    Dim sStart As Single
    lEnd = rg.End
    
    With rg.Find
        .ClearFormatting
        .MatchWholeWord = True
        .MatchCase = True
        .Wrap = wdFindStop
        .Forward = True
        .Text = sFrom
        '.Replacement.Text = sTo
        .Execute Wrap:=wdFindStop 'Replace:=wdReplaceOne
        If InStr(sTo, sFrom) Then
            sStart = Timer
            While .Found
                If Timer - sStart > 30 Then
                    If MsgBox("30 seconds passed, continue?", vbYesNo) = vbNo Then
                        Exit Function
                    End If
                End If
                lTmpStart = rg.Start
                lTmpEnd = rg.End
                rg.SetRange rg.Start - InStr(sTo, sFrom) + 1, rg.End + Len(sTo) - InStr(sTo, sFrom) + 1 - Len(sFrom)
                If rg.Text = sTo Then
                    rg.SetRange rg.End, rg.End 'lEnd
                    ReplaceString rg, sFrom, sTo
                Else
                    rg.SetRange lTmpStart, lTmpEnd
                    GoTo Finish
                End If
            Wend
        Else
            If .Found Then
Finish:         rg.Text = sTo
                rg.HighlightColorIndex = wdBrightGreen
                ReplaceString = True
                Exit Function
            End If
        End If
    End With
End Function
