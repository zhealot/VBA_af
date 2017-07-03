Attribute VB_Name = "Module1"
'constants for control positions
Public Const WIDTH_SECTION = 120        'width for checkbox in SECTIONS
Public Const HEIGHT_SECTION = 21        'height for checkbox in SECTIONS
Public Const TOP_GAP = 24               'gap betweeen two rows of control
Public Const LEFT_COLUMN_1 = 6          'check box in column 1 left distance
Public Const LEFT_COLUMN_2 = 135        'check box in column 2 left distance
Public Const TAG_SECTION = "SECTION"    'tag name for controls in SECTIONS
Public Const FONT_SIZE = 10             'font size for checkbox
Public Const FONT_NAME = "Segoe UI"     'font name for checkbox
Public Const SECTION_NAMES = "1 Policy,2 H&S Tasking,3 Emergency,4 Worker,5 Hazard,6 First Aid,7 Contractor,8 Toolbox,9 Machinery/Tools,10 Vehicle"
'constants for label
Public Const TOP_GAP_LABEL = 18         'gap between labels
Public Const LEFT_LABEL = 12            'left distance for labels
Public Const FONT_SIZE_LABEL = 12       'label font size
Public Const WIDTH_LABEL = 240          'label width
Public Const HEIGHT_LABEL = 16          'label height
Public Const DATE_FORMAT = "dd-mmmm-yyyy" 'date format
Public Const NUMBER_LIST_STYLE = "List Number"  'style for number list
Public Const NUMBER_LIST_2_STYLE = "List Number 2" 'style for number list 2
'collection to hold controls
Public cbSections() As New cbHandler       'checkboxes in frame SECTION
'Public cbTemplates() As CheckBox       'checkboxes in frame TEMPLATES
Public cbSelected() As New cbTemplates           'templates which have been selected
Public Blocks() As New Block

Sub CallBack(control As IRibbonControl)
    fmMain.Show
End Sub

'scan folder and add content of each docx as a building bolcks
Sub AddBuildingBlock(control As IRibbonControl)
    MsgBox "Please select a folder that contains templates."
    Application.ScreenUpdating = False
    Dim s As Long: s = Timer
    Dim doc As Document
    Dim Fld As String   'folder name
    Dim File As String  'file name
    Dim Ext As String: Ext = "*.docx"   'file extension
    Dim bbName As String   'name for building block
    Dim posNumeric As Integer  'position of numeric char in file name
    Dim rg As Range
        
    'choose templates folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .InitialFileName = ThisDocument.Path & "\"
        .ButtonName = "OK"
        If .Show = -1 Then
            Fld = .SelectedItems(1) & "\"
        End If
    End With
    
    'user pressed 'Cancel'
    If Fld = "" Then End
    
    'delete existing building blocks
    While ThisDocument.AttachedTemplate.BuildingBlockEntries.Count > 0
        ThisDocument.AttachedTemplate.BuildingBlockEntries(1).Delete
    Wend
    
    'look if folder contains valid file (has a number in filename)
    File = Dir(Fld & Ext, vbNormal)
    While File <> ""
        'check if file name starts with number
        If IsNumeric(Left(File, 1)) Then
            bbName = Left(File, Len(File) - 5)
        'or has number in the middle
        Else
            For i = 1 To Len(File)
                If IsNumeric(Mid(File, i, 1)) Then
                    bbName = Mid(File, i, Len(File) - i - 4)
                    Exit For
                End If
            Next i
        End If
        
        'Debug.Print bbName
        Set doc = Documents.Open(Fld & File, ConfirmConversions:=False, ReadOnly:=True, Visible:=False)
        'add section breaks before add as Building Blocks
        Set rg = doc.Content
        rg.Collapse wdCollapseStart
        rg.InsertBreak wdSectionBreakOddPage 'start new seciton on odd page
        Set rg = doc.Content
        rg.Collapse wdCollapseEnd
        rg.InsertBreak wdSectionBreakNextPage
        'in order to keeep the original formatting, copy and paste to thisdocument before adding it to BB
        ThisDocument.Content.Delete
        Set rg = doc.Content
        rg.Copy
        'work around: some templates need to be added to BB directly
        Dim PasteMethod As Integer
        PasteMethod = wdUseDestinationStylesRecovery
'        Select Case LCase(Left(bbName, 2))
'            Case "1.", "7b", "5a", "5e"
'                PasteMethod = wdUseDestinationStylesRecovery
'            Case Else
'                PasteMethod = wdFormatOriginalFormatting
'        End Select
         ThisDocument.Content.PasteAndFormat PasteMethod
        Set rg = ThisDocument.Content
       'add docx content to building blocks
       'note: BB name can not be more than 32 chars
        ThisDocument.AttachedTemplate.BuildingBlockEntries.Add Name:=Left(bbName, 32), _
                                                              Type:=wdTypeQuickParts, _
                                                              Category:="General", _
                                                              Description:=bbName, _
                                                              Range:=rg, _
                                                              InsertOptions:=wdInsertContent
        doc.Close False
        DoEvents
        File = Dir  'get next file
    Wend   'end While File<>""
    ThisDocument.Content.Delete
    'ThisDocument.Save
    ClearClipBoard
    Application.ScreenUpdating = True
    Debug.Print Timer - s
    MsgBox "Templates Updated!"
End Sub

'button visibility in ribbon
Sub ReturnVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IIf(ActiveDocument.Type = wdTypeTemplate, True, False)
End Sub

'read custom property
Function ReadCP(tb As MSForms.TextBox)
    On Error Resume Next
    ReadCP = ActiveDocument.CustomDocumentProperties(tb.Name)
End Function

'write custom property
Function WriteCP(tb As MSForms.TextBox)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(tb.Name) = tb.Text
End Function

'check textbox
Function CheckTB(fm As MSForms.Frame) As Boolean
    Dim v As MSForms.control
    For Each v In fm.Controls
        If TypeName(v) = "TextBox" Then
            If Trim(v.Text) = "" Then
                v.SetFocus
                CheckTB = False
                Exit Function
            End If
        End If
    Next
    CheckTB = True
End Function

'reaplace picture in header's contentcontrol
Function ReplacePicInHeader(hf As HeaderFooter)
    Dim SCT As Section
    Dim CC As ContentControl
    If hf.Range.ShapeRange.Count > 0 Then
        For i = 1 To hf.Range.ShapeRange.Count
            If hf.Range.ShapeRange(i).Type = msoTextBox Then
                If hf.Range.ShapeRange(i).TextFrame.TextRange.ContentControls.Count > 0 Then
                    Set CC = hf.Range.ShapeRange(i).TextFrame.TextRange.ContentControls(1)
                    If CC.Range.InlineShapes.Count > 0 Then
                        Set rg = CC.Range
                        CC.Range.InlineShapes(1).Delete         'in case been set before
                        Set rg = CC.Range
                        rg.InlineShapes.AddPicture fmMain.imgLogo.Tag, False, True, rg
                    End If
                End If
            End If
        Next i
        If hf.Range.ShapeRange(1).TextFrame.HasText <> 0 Then
        End If
    End If
End Function

'parse file name to make: '1a' comes before '10a', etc
Function ParseFileName(s As String)
    If Len(s) < 2 Then
        ParseFileName = s
    ElseIf IsNumeric(Left(s, 1)) Then
        If Mid(s, 2, 1) = "." Or Not IsNumeric(Left(s, 2)) Then 'make '1.' to '01.', '1a' to '01a'
            ParseFileName = "0" & s
        ElseIf IsNumeric(Left(s, 2)) Then
            ParseFileName = s
        End If
    End If
End Function

Public Function ClearClipBoard()
    Dim oData   As New DataObject 'object to use the clipboard
    oData.SetText Text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
End Function
 
Sub steste()
    Dim rg As Range
    Set rg = Selection.Range.Paragraphs(1).Range
    'rg.Collapse
    'rg.ListFormat.ApplyListTemplate ListGalleries(wdOutlineNumberGallery).ListTemplates(7), True
    
    rg.Style = "List Number"
End Sub

Function ParaIndentment(rg As Range)
    With rg.ParagraphFormat
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
    End With
End Function
