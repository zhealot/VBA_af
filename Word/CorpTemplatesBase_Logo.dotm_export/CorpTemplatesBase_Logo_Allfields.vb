Attribute VB_Name = "Allfields"
Option Explicit
Dim oWithEvent As New EventsClass
'module to hold ribbon button code
Const Style1 = "Instructional Text"
Const Style2 = "Instructional Text Bullets"
Public Const DOCUMENTPROPERTY = "Category" 'document property name to hold business group
Public sBG As String 'short business group name
Public docA As Document


Public Sub AutoExec()
    Set oWithEvent.oWdApp = Word.Application
End Sub

Public Sub OnLoad(control As IRibbonControl)
    MsgBox "on open"
    Call AutoExec
End Sub

Public Sub ShowDIP(control As IRibbonControl)
    On Error Resume Next
    If Application.DisplayDocumentInformationPanel = False Then
        Application.DisplayDocumentInformationPanel = True
    Else
        Application.DisplayDocumentInformationPanel = False
    End If
End Sub

Public Sub RemoveInstructions(control As IRibbonControl)
'Search for paragraphs styled "Instructions" and delete them
    Dim rng As Range
    Dim boolFound As Boolean
    Dim boolSmartCutPaste As Boolean
    Dim MsgText As String
    
    'Smart cut & paste' setting must be false so that deleting the last paragraph
    'in a table cell doesn't change the style
    
    boolSmartCutPaste = Options.PasteSmartCutPaste
    Options.PasteSmartCutPaste = False
    
    On Error GoTo Bye
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(Style1)
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Execute
        If rng.End = ActiveDocument.Content.End Then GoTo NextStyle   'not whole document!
        While .Found
            boolFound = True
            rng.Delete
            If rng.Information(wdWithInTable) Then
                'if at end of cell, no paragraph to delete
                If rng.Start = rng.Cells(1).Range.End - 1 Then
                    On Error Resume Next
                    rng.Style = "TinyFont"
                    On Error GoTo Bye
                End If
            End If
            rng.SetRange Start:=rng.Start, End:=ActiveDocument.Range.End
            .Execute
        Wend
    End With
NextStyle:
    'repeat for different style
    Set rng = ActiveDocument.Content
    With rng.Find
        .Style = ActiveDocument.Styles(Style2)
        .Execute
        If rng.End = ActiveDocument.Content.End Then GoTo Bye   'not whole document!
        While .Found
            boolFound = True
            rng.Delete
            If rng.Information(wdWithInTable) Then
                'if at end of cell, no paragraph to delete
                If rng.Start = rng.Cells(1).Range.End - 1 Then
                    On Error Resume Next
                    rng.Style = "TinyFont"
                    On Error GoTo Bye
                End If
            End If
            rng.SetRange Start:=rng.Start, End:=ActiveDocument.Range.End
            .Execute
        Wend
    End With
    
Bye:
    Err.Clear
    'restore 'Smart cut & paste' setting
    Options.PasteSmartCutPaste = boolSmartCutPaste
    
    If boolFound Then
        MsgText = "All instructions have been removed."
    Else
        MsgText = "No instructions were found."
    End If
    'MsgBox MsgText, vbInformation, MED
End Sub

Public Sub Logo(control As IRibbonControl)
    fmLogo.Show
End Sub

Public Function OBClick(ob As control)
    sBG = ""
    Dim ctl As control
    For Each ctl In fmLogo.frmImage.Controls
        If Right(ctl.Name, Len(ctl.Name) - 3) = Right(ob.Name, Len(ob.Name) - 2) Then
            sBG = Right(ob.Name, Len(ob.Name) - 2)
            ctl.Visible = True
        Else
            ctl.Visible = False
        End If
    Next ctl
End Function

Public Function SetLogo(docA As Document)
    'Dim docA As Document
    Dim rg As Range
    Dim rgTmp As Range
    Dim rgCurrent As Range
    Dim oApp As Word.Application
    Dim tmp As Template
    Set oApp = Word.Application
    oApp.ScreenUpdating = False
    'check if Business Group is set
    If sBG = "" Then
        sBG = ReadBG
    End If
    If sBG = "" Then
        Exit Function
    End If
    'set range to whole current page
    Set rg = docA.Content
    rg.Collapse wdCollapseStart
    Dim iSec As Integer
    Dim iHdr As Integer
    Dim SeekView As Long
    Dim docTmp As Template
    Dim spOri As Shape  'original shape
    Dim spNew As Shape  'new shape
    For Each docTmp In oApp.Templates
        If docTmp.Name = ThisDocument.Name Then
            For iSec = 1 To docA.Sections.Count
                For iHdr = 1 To 3   'check header in primary/evenpage/firstpage
                    If docA.Sections(iSec).Headers(iHdr).Exists Then
                        'docA.ActiveWindow.View.SeekView = wdSeekCurrentPageHeader
                        Set rg = docA.Sections(iSec).Headers(iHdr).Range
                        If rg.ShapeRange.Count > 0 Then
                            Set spOri = rg.ShapeRange(1)
                            rg.Collapse wdCollapseEnd
                            Set rgTmp = docTmp.BuildingBlockEntries(sBG).Insert(rg)
                            Set spNew = rgTmp.ShapeRange(1)
                            'spNew.RelativeHorizontalPosition = spOri.RelativeHorizontalPosition
                            'spNew.RelativeHorizontalSize = spOri.RelativeHorizontalSize
                            spNew.RelativeVerticalPosition = spOri.RelativeVerticalPosition
                            spNew.RelativeVerticalSize = spOri.RelativeVerticalSize
                            spNew.LockAspectRatio = spOri.LockAspectRatio
                            spNew.Height = spOri.Height
                            'spNew.Width = spOri.Width
                            spNew.WrapFormat.Type = spOri.WrapFormat.Type
                            spNew.Left = wdShapeRight 'spOri.Left
                            spNew.Top = spOri.Top
                            spOri.Delete
                        End If
                    End If
                Next iHdr
            Next iSec
            Exit For
        End If
    Next docTmp
    For Each docTmp In oApp.Templates
        If LCase(docTmp) = "normal.dotm" Then
            docTmp.BuiltInDocumentProperties(DOCUMENTPROPERTY) = sBG
            docTmp.Save
            Exit For
        End If
    Next docTmp
    oApp.ScreenUpdating = True
End Function

Public Function ReadBG() As String
    If LCase(ActiveDocument.Name) = "normal.dotm" Then
        Exit Function
    End If
    Dim tp As Template
    For Each tp In Application.Templates
        If LCase(tp.Name) = "normal.dotm" Then
            ReadBG = tp.BuiltInDocumentProperties(DOCUMENTPROPERTY)
        End If
    Next tp
End Function
