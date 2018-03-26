Attribute VB_Name = "Module1"
'----------------------------------   -------------------------------------------
' Developed for Cato Bolam
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             March 2018
' Description:      Check document for error and highlight etc, add cover page & address in footer then export to PDF
'-----------------------------------------------------------------------------
Public Const HIGHLIGHTFOUNDTEXT = "Highlight found."
Public Const SEARCHFINISHTEXT = "We have checked the Document for highlights and fields." & vbNewLine & "All seems OK now for you to add a cover and save as PDF." & vbNewLine & "Now click Export PDF on the Report Finalisation Ribbon."
Public Const LETTERHEADER = "LetterHeader"
Public Const SIGNINGCOVER1 = "ReportCoverFormal"
Public Const SIGNINGCOVER2 = "ReportCoverHeader"
Public sCover As String
Public sAddress As String
Public sSigningCover As String
Public docA As Document

'Callback for customButton onAction
Sub Check(control As IRibbonControl)
    'Dim docA As Document
    Set docA = ActiveDocument
    Dim rg As Range
'    Set rg = docA.Range
'    If SearchRange(rg) Then
'        MsgBox HIGHLIGHTFOUNDTEXT
'        Exit Sub
'    End If

    'look in headers & footers
'    Dim sec As Section
'    For Each sec In docA.Sections
'        For i = 1 To 3  'even/first/primary
'            If sec.Headers(i).Exists Then
'                If SearchRange(sec.Headers(i).Range) Then
'                    MsgBox HIGHLIGHTFOUNDTEXT
'                    Exit Sub
'                End If
'            End If
'            If sec.Footers(i).Exists Then
'                If SearchRange(sec.Footers(i).Range) Then
'                    MsgBox HIGHLIGHTFOUNDTEXT
'                    Exit Sub
'                End If
'            End If
'        Next i
'    Next sec
    For Each rg In docA.StoryRanges
        If SearchRange(rg) Then
            MsgBox HIGHLIGHTFOUNDTEXT
            Exit Sub
        End If
        Do While Not rg.NextStoryRange Is Nothing
            Set rg = rg.NextStoryRange
            If SearchRange(rg) Then
                MsgBox HIGHLIGHTFOUNDTEXT
                Exit Sub
            End If
        Loop
    Next rg

    'update fields
    docA.Content.Fields.Update
    For Each rg In docA.StoryRanges
        rg.Fields.Update
    Next rg
    'if nothing foud, message to notify user
    MsgBox SEARCHFINISHTEXT
End Sub

'Callback for toPDF onAction
Sub ExportPDF(control As IRibbonControl)
    '###TBD: check for multiple documents
    If Not docA Is Nothing Then
        If ActiveDocument.FullName <> docA.FullName Then
            Unload UserForm1
        End If
    End If
    Set docA = ActiveDocument
    docA.Activate
    docA.Windows(1).Activate
    UserForm1.Show
End Sub

Sub SpellingCheckk(control As IRibbonControl)
    If ActiveDocument.SpellingErrors.Count > 0 Then
        ActiveDocument.CheckSpelling
    Else
        MsgBox "No spelling error found."
    End If
End Sub

'insert header image for Appendixheader
Sub AppendixHeader(control As IRibbonControl)
    Dim oApp As Application
    Set oApp = Application
    Dim tmp As Template
    
    For Each tmp In oApp.Templates
        On Error Resume Next
        If tmp.Name = ThisDocument.Name Then
            Dim rg As Range
            Set rg = Selection.Range
            rg.Collapse wdCollapseStart
            Set rg = tmp.BuildingBlockEntries(LETTERHEADER).Insert(rg)
            'adjust cover page image
            Dim sp As Shape
            Set sp = rg.ShapeRange(1)
            If Not sp Is Nothing Then
                With sp
                    .WrapFormat.Type = wdWrapBehind
                    .LockAspectRatio = msoFalse
                    .Top = 0
                    .Left = 0
                    .Width = docA.Sections(rg.Information(wdActiveEndSectionNumber)).PageSetup.PageWidth
                End With
            End If
            Exit For
        End If
    Next tmp
End Sub

Public Function ChooseCover(ob As OptionButton)
    Dim ctrl As control
    For Each ctrl In UserForm1.Controls
        If Left(ctrl.Name, 3) = "img" Then
            ctrl.Visible = IIf(Right(ctrl.Name, Len(ctrl.Name) - 3) = Right(ob.Name, Len(ob.Name) - 2), True, False)
        End If
    Next
    sCover = Right(ob.Name, Len(ob.Name) - 2)
    ControlSwitch UserForm1.fmFooter, IIf(sCover = LETTERHEADER, True, False)
End Function

Public Function ChooseSigningCover(ob As OptionButton)
    sSigningCover = Right(ob.Name, Len(ob.Name) - 2)
    'name casting
    Select Case sSigningCover
    Case "SigningCover1"
        sSigningCover = SIGNINGCOVER1
    Case "SigningCover2"
        sSigningCover = SIGNINGCOVER2
    Case Else
        sSigningCover = ""
    End Select
End Function

Public Function ChooseFooter(ob As OptionButton)
    sAddress = Right(ob.Name, Len(ob.Name) - 2)
End Function

Public Function ControlSwitch(ctr As control, YesNo As Boolean)
    ctr.Enabled = YesNo
    For Each c In ctr.Controls
        c.Enabled = YesNo
    Next c
End Function

Sub test()
    If ActiveDocument.SpellingErrors.Count > 0 Then
        ActiveDocument.CheckSpelling
    Else
        MsgBox "No spelling error found."
    End If
End Sub

Public Function SearchRange(rg As Range) As Boolean
    With rg.Find
        .Highlight = True
        .Execute
        If .Found Then
            'set view to avoiding switched to draft(normal) view
            If docA.ActiveWindow.View <> wdNormalView Then
                Select Case rg.Information(wdHeaderFooterType)
                Case "0", "1", "4"  'header view
                    ActiveWindow.View.SeekView = wdSeekCurrentPageHeader
                Case "2", "3", "5": 'footer view
                    ActiveWindow.View.SeekView = wdSeekCurrentPageFooter
                End Select
            End If
            rg.Select
        End If
        SearchRange = .Found
    End With
End Function
