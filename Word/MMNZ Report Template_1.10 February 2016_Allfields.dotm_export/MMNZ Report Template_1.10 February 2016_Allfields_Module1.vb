Attribute VB_Name = "Module1"
Option Explicit

Sub UpdateRepProp(control As IRibbonControl)

'   Load Report Properties Form
    Load frmDocProp
    
'   Populate Document Properties if Already Existing
    If ActiveDocument.CustomDocumentProperties("Client") <> "" Then
        frmDocProp.txtClient.Text = ActiveDocument.CustomDocumentProperties("Client")
    End If

    If ActiveDocument.CustomDocumentProperties("Report Date") <> "" Then
        frmDocProp.txtReportDate.Text = ActiveDocument.CustomDocumentProperties("Report Date")
    End If
    
    If ActiveDocument.CustomDocumentProperties("Report Title") <> "" Then
        frmDocProp.txtReportTitle.Text = ActiveDocument.CustomDocumentProperties("Report Title")
    End If
    
    If ActiveDocument.CustomDocumentProperties("Header Summary") <> "" Then
        frmDocProp.txtHeaderSummary.Text = ActiveDocument.CustomDocumentProperties("Header Summary")
    End If
    
    If ActiveDocument.CustomDocumentProperties("Report No") <> "" Then
        frmDocProp.txtReportNo.Text = ActiveDocument.CustomDocumentProperties("Report No")
    End If
    
    If ActiveDocument.CustomDocumentProperties("Issue") <> "" Then
        frmDocProp.cboIssue.Text = ActiveDocument.CustomDocumentProperties("Issue")
    End If
    
    If ActiveDocument.CustomDocumentProperties("Report Security") <> "" Then
        frmDocProp.cboSecurity.Text = ActiveDocument.CustomDocumentProperties("Report Security")
    End If
        
    If ActiveDocument.CustomDocumentProperties("Checked") <> "" Then
        frmDocProp.txtChecked.Text = ActiveDocument.CustomDocumentProperties("Checked")
    End If
    
    frmDocProp.Show
    
End Sub

Sub InsertMainSec(control As IRibbonControl)

Dim cntSections, thisSection

Application.ScreenUpdating = False

'    ActiveDocument.Unprotect Password:="w1n1"

    
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    
    With Selection.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    End With

    cntSections = ActiveDocument.Sections.Count
    thisSection = Selection.Information(wdActiveEndSectionNumber)
    
    ActiveDocument.Sections(thisSection).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
    ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
          
    ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).Range.Select
    
    Selection.Font.Size = 10
    
'   Add Footer Info
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DOCPROPERTY  Client ", PreserveFormatting:=True

    Selection.ParagraphFormat.TabStops.ClearAll

     Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(16), _
        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces

    Selection.TypeText Text:=vbTab

'   Insert & Format Page Numering
    With Selection
        .TypeText Text:="Page "
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "PAGE ", PreserveFormatting:=True
        .TypeText Text:=" of "
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "SECTIONPAGES ", PreserveFormatting:=True
    End With

    With ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).PageNumbers
        .NumberStyle = wdPageNumberStyleArabic
        .RestartNumberingAtSection = True
        .StartingNumber = 1
    End With

'   Return to Main Window
    ActiveWindow.Panes(2).Close
    
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
'   Update Fields
    Selection.WholeStory
    Selection.Fields.Update
    


'   Reprotect Document
'    Selection.WholeStory
'    Selection.Editors.Add wdEditorEveryone
'    ActiveDocument.Protect Password:="w1n1", NoReset:=False, Type:= _
'        wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False
        
'    ActiveDocument.FormattingShowFilter = wdShowFilterStylesAvailable
        
    Selection.GoTo wdGoToBookmark, , , "\EndOfDoc"
    
Application.ScreenUpdating = True



End Sub

Sub InsertAnnexSec(control As IRibbonControl)

Dim cntSections, thisSection, x

Application.ScreenUpdating = False

'    ActiveDocument.Unprotect Password:="w1n1"
  
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    
    cntSections = ActiveDocument.Sections.Count
    thisSection = Selection.Information(wdActiveEndSectionNumber)
'    x = MsgBox("Count of sections = " + CStr(cntSections), vbOKOnly)

    ActiveDocument.Sections(thisSection).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
    ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
    
    With Selection.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    End With
    
    Selection.Style = ActiveDocument.Styles("Heading 9")
    
    ActiveDocument.Sections(thisSection).Headers(wdHeaderFooterPrimary).Range.Select

'   Change Header Tabs
    
    Selection.ParagraphFormat.TabStops.ClearAll
    
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(8), _
        Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces

     Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(16), _
        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        
'   Return to Main Window
    ActiveWindow.Panes(2).Close
    
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If

    
    ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).Range.Select
    
'   Add Footer Info
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DOCPROPERTY  Client ", PreserveFormatting:=True

    Selection.ParagraphFormat.TabStops.ClearAll

     Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(16), _
        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces

    Selection.TypeText Text:=vbTab
    
'   Insert & Format Page Numering
    With Selection
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "PAGE ", PreserveFormatting:=True
    End With

    With ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).PageNumbers
        .NumberStyle = wdPageNumberStyleArabic
        .RestartNumberingAtSection = True
        .StartingNumber = 1
    End With

    With ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).PageNumbers
        .NumberStyle = wdPageNumberStyleArabic
        .IncludeChapterNumber = True
        .HeadingLevelForChapter = 8 ' Heading 9, Styles 0-8
        .ChapterPageSeparator = wdSeparatorHyphen
        .RestartNumberingAtSection = True
        .StartingNumber = 1
    End With
    
    ActiveWindow.Panes(2).Close
    
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
'   Reprotect Document
'    Selection.WholeStory
'    Selection.Editors.Add wdEditorEveryone
'    ActiveDocument.Protect Password:="w1n1", NoReset:=False, Type:= _
'        wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False
        
    Selection.GoTo wdGoToBookmark, , , "\EndOfDoc"
        
Application.ScreenUpdating = True

End Sub

Sub InsertAnnexA3Sec(control As IRibbonControl)

Dim cntSections, thisSection, x

Application.ScreenUpdating = False

'    ActiveDocument.Unprotect Password:="w1n1"
    
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    
    cntSections = ActiveDocument.Sections.Count
    thisSection = Selection.Information(wdActiveEndSectionNumber)
'    x = MsgBox(thisSection)

    ActiveDocument.Sections(thisSection).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
    ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
    
    With Selection.PageSetup
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(42)
        .PageHeight = CentimetersToPoints(29.7)
    End With
    
'### apply
    Selection.Style = ActiveDocument.Styles("Heading 9")
    
    With ActiveDocument.Sections(thisSection).Headers(wdHeaderFooterPrimary)
        .Range.Select
        .Range.Style = "A3Header"
    End With
    
    ActiveWindow.Panes(2).Close
    
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
    With ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary)
        .Range.Select
        .Range.Style = "A3Footer"
    End With
    
    '   Add Footer Info
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DOCPROPERTY  Client ", PreserveFormatting:=True
        
    Selection.ParagraphFormat.TabStops.ClearAll

         Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(37), _
        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        
    Selection.TypeText Text:=vbTab
    
'   Insert & Format Page Numering
    With Selection
        .TypeText Text:="Page "
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "PAGE ", PreserveFormatting:=True
'        .TypeText Text:=" of "
'        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
'            "SECTIONPAGES ", PreserveFormatting:=True
    End With

    With ActiveDocument.Sections(thisSection).Footers(wdHeaderFooterPrimary).PageNumbers
        .NumberStyle = wdPageNumberStyleArabic
        .HeadingLevelForChapter = 0
        .IncludeChapterNumber = True
        .HeadingLevelForChapter = 8 ' Heading 9, Styles 0-8
        .ChapterPageSeparator = wdSeparatorHyphen
        .RestartNumberingAtSection = True
        .StartingNumber = 1
    End With
    
    ActiveWindow.Panes(2).Close
    
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If

'   Reprotect Document
'    Selection.WholeStory
'    Selection.Editors.Add wdEditorEveryone
'    ActiveDocument.Protect Password:="w1n1", NoReset:=False, Type:= _
'        wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False
        
    Selection.GoTo wdGoToBookmark, , , "\EndOfDoc"
    
Application.ScreenUpdating = True

End Sub

'### intercept build in function InsertCaption, put caption in textbox and formatting, tao@allfields.co.nz, 29/10/2015
Sub InsertCaption()
    If Dialogs(wdDialogInsertCaption).Show = -1 Then    'clicked OK
        Application.ScreenUpdating = False
        On Error Resume Next
        Dim rG As Range
        Set rG = Selection.Paragraphs.First.Range
        rG.Cut
        rG.InsertParagraphBefore
        rG.Collapse
        Dim txtBox As Shape
        Set txtBox = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, rG.Information(wdHorizontalPositionRelativeToPage), rG.Information(wdVerticalPositionRelativeToPage), 1, 20)
        With txtBox
            .TextFrame.TextRange.PasteAndFormat wdFormatOriginalFormatting
            .TextFrame.TextRange.Font.Bold = True   ' set fond to bold
            .TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0    'spaceing before
            .TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0 'spaing after
            .TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphCenter 'text alignment: centred
            .TextFrame.TextRange.ParagraphFormat.LineSpacing = LinesToPoints(1) 'line spacing
            .TextFrame.WordWrap = False 'resize text box to fit text
            .Line.Visible = msoFalse    'no boundary lines
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage  'horizontal alignment: centred
            .Left = wdShapeCenter   'horizontal alignment: centred
            .LayoutInCell = False   'uncheck layout in table cell
        End With
        On Error GoTo 0
        Application.ScreenUpdating = True
    End If
End Sub

Sub FileSaveAs()
'
' FileSaveAs Macro
' Saves a copy of the document in a separate file
'
'MsgBox "hadf"
'    Dialogs(wdDialogFileSaveAs).Show
'    MsgBox "hah"
    Call UpdateHeadersFooters
End Sub

Sub InsertImage(control As IRibbonControl)

Dim dlgDoc As Dialog
Dim cmdInsert As String
Dim strImgPath As String
Dim strImgName As String

Set dlgDoc = Dialogs(wdDialogInsertPicture)

    With dlgDoc
        .Name = "*.jpg"
        cmdInsert = .Show
        strImgPath = .Name
        
    End With
    
'   Strip off the path and file ext. to get the file name

    strImgName = Right(strImgPath, Len(strImgPath) - InStrRev(strImgPath, "\"))
    strImgName = Left([strImgName], InStrRev([strImgName], ".") - 1)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

    Selection.ParagraphFormat.KeepWithNext = True
    
    Selection.TypeParagraph

'   Insert strFileName as caption
'### bypass error when instering into footer, 29/10/2015, tao@allfields.co.nz
    On Error Resume Next
    If cmdInsert = -1 Then
        Selection.InsertCaption Label:="Figure", _
        TitleAutoText:="", Title:=": " & strImgName, _
        Position:=wdCaptionPositionBelow
    End If
    
'   Insert an image border

    Call ImgBorder
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
End Sub

Sub ImgBorder()

'   Insert a .5 border around images

Dim objInlineImg As InlineShape

    For Each objInlineImg In ActiveDocument.InlineShapes
    
        With objInlineImg
                If .Width > CentimetersToPoints(10) Then
                    .Borders.OutsideLineStyle = wdLineStyleSingle
                    .Borders.OutsideLineWidth = wdLineWidth050pt
                    '.Width = CentimetersToPoints(15)
                End If
        End With
    
    Next

End Sub

Sub UpdateRefs(control As IRibbonControl)

Application.ScreenUpdating = False

    ActiveDocument.TablesOfContents(1).Update
    ActiveDocument.TablesOfContents(2).Update
    ActiveDocument.TablesOfFigures(1).Update
    ActiveDocument.TablesOfFigures(2).Update
    
Application.ScreenUpdating = False

End Sub

'   update all fields codes in heaters and footers, tao@allfields.co.nz, 9/11/2015
Public Sub UpdateHeadersFooters()
    Dim sSection As Section
    Dim HF As HeaderFooter
    
    For Each sSection In ActiveDocument.Sections
        For Each HF In sSection.Headers
            HF.Range.Fields.Update
        Next HF
        For Each HF In sSection.Footers
            HF.Range.Fields.Update
        Next HF
    Next sSection
End Sub

' insert A3 page, with continuing page number and footer & header elements, tao@allfields.co.nz, 6/11/2015
Sub InsertA3Page(control As IRibbonControl)
    
    Dim BaseSection As Integer
    Dim TmpSection As Integer
    
    Application.ScreenUpdating = False
    BaseSection = Selection.Information(wdActiveEndSectionNumber)
    Dim IsLastPage As Boolean
    
    IsLastPage = IIf(Selection.Information(wdActiveEndPageNumber) = Selection.Information(wdNumberOfPagesInDocument), True, False)
        
    On Error Resume Next
    With Selection
        .Bookmarks("\page").Range.Select
        .Collapse wdCollapseEnd
        TmpSection = .Information(wdActiveEndSectionNumber)
        .MoveLeft wdCharacter, 1
        .InsertBreak Type:=wdSectionBreakNextPage
        If Not IsLastPage And TmpSection = BaseSection Then
            .InsertBreak Type:=wdSectionBreakNextPage
            .GoToPrevious (wdGoToPage)
        End If
    End With
    
    With Selection.PageSetup
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(42)
        .PageHeight = CentimetersToPoints(29.7)
    End With
    
    If Not IsLastPage Then
        With ActiveDocument.Sections(BaseSection + 2).Headers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .PageNumbers.RestartNumberingAtSection = False
        End With
        With ActiveDocument.Sections(BaseSection + 2).Footers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .PageNumbers.RestartNumberingAtSection = False
        End With
    End If
         
    With ActiveDocument.Sections(BaseSection + 1).Headers(wdHeaderFooterPrimary)
        .LinkToPrevious = False
        .Range.Select
        .Range.Style = "A3Header"
    End With
    
    ActiveWindow.Panes(2).Close
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
    With ActiveDocument.Sections(BaseSection + 1).Footers(wdHeaderFooterPrimary)
        .LinkToPrevious = False
        .PageNumbers.RestartNumberingAtSection = False
        If .Range.Tables.Count = 0 Then
            .Range.ParagraphFormat.TabStops.ClearAll
            .Range.ParagraphFormat.TabStops.Add 0, wdAlignTabRight, wdTabLeaderSpaces
        Else
            .Range.Tables(1).AutoFitBehavior wdAutoFitWindow
        End If
        
    End With

    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
    Application.ScreenUpdating = True
End Sub
