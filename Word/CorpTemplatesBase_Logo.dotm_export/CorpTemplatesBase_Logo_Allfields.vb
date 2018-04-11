Attribute VB_Name = "Allfields"
Option Explicit
Dim oWithEvent As New EventsClass
'module to hold ribbon button code
Const Style1 = "Instructional Text"
Const Style2 = "Instructional Text Bullets"
Const SERVERPROPERTY = "Portfolio"  'property to read from server
Const HEADER_TOP_TO_PAGE = 1.2
Const HEADER_LEFT_TO_PAGE = 11.3
Const FOOTER_TOP_TO_PAGE = 27
Const FOOTER_LEFT_TO_PAGE = 1.2
Const HEADER_LOGO_WIDTH = 8.6
Const FOOTER_LOGO_WIDTH = 6.2
Public Const DOCUMENTPROPERTY = "Category" 'document property name to hold business group
Public sBG As String 'short business group name
Public doca As Document



Public Sub AutoExec()
    Set oWithEvent.oWdApp = Word.Application
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

Public Function SetLogo(doca As Document, Optional sPty As String = "")
'sPty is used for bath convert document based on document server property
    Application.ScreenUpdating = False
    'Dim docA As Document
    Dim rg As Range
    Dim rgTmp As Range
    Dim rgCurrent As Range
    Dim oApp As Word.Application
    Dim tmp As Template
    Set oApp = Word.Application
    oApp.ScreenUpdating = False
    'check if Business Group is set
    If sPty <> "" Then
        sBG = ReadServerProperty(doca, SERVERPROPERTY)
    End If
    If sBG = "" Then
        Exit Function
    End If
    'set range to whole current page
    Set rg = doca.Content
    rg.Collapse wdCollapseStart
    Dim iSec As Integer
    Dim iHdr As Integer
    Dim SeekView As Long
    Dim docTmp As Template
    Dim spOri As Shape  'original shape
    Dim spNew As Shape  'new shape
    For Each docTmp In oApp.Templates
        If docTmp.Name = ThisDocument.Name Then
            For iSec = 1 To doca.Sections.Count
                For iHdr = 1 To 3   'check header in primary/evenpage/firstpage
                    If doca.Sections(iSec).Headers(iHdr).Exists Then
                        'docA.ActiveWindow.View.SeekView = wdSeekCurrentPageHeader
                        Set rg = doca.Sections(iSec).Headers(iHdr).Range
                        If rg.ShapeRange.Count > 0 Then
                            Set spOri = rg.ShapeRange(1)
                            spOri.Delete
                            rg.Collapse wdCollapseEnd
                            Set rgTmp = docTmp.BuildingBlockEntries(sBG).Insert(rg)
                            Set spNew = rgTmp.ShapeRange(1)
                            spNew.LockAspectRatio = msoTrue
                            spNew.Width = CentimetersToPoints(HEADER_LOGO_WIDTH)
                            spNew.WrapFormat.Type = wdWrapBehind
                            spNew.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                            spNew.Left = CentimetersToPoints(HEADER_LEFT_TO_PAGE)
                            spNew.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                            spNew.TOP = CentimetersToPoints(HEADER_TOP_TO_PAGE)
                        End If
                    End If
                    'deal with the footer part
                    If doca.Sections(iSec).Footers(iHdr).Exists Then
                        Set rg = doca.Sections(iSec).Footers(iHdr).Range
                        'delete existing logo
                        If rg.ShapeRange.Count > 0 Then
                            Dim i As Integer
                            For i = rg.ShapeRange.Count To 1 Step -1
                                If rg.ShapeRange(i).Title = "MPI" Then
                                    rg.ShapeRange(i).Delete
                                End If
                            Next i
                        End If
                        Select Case sBG
                        Case "Bio", "Fis", "NZF"    'for those groups, insert MPI logo
                            If rg.ShapeRange.Count > 0 Then
                                If rg.ShapeRange(1).Title = "mpi" Then
                                    rg.ShapeRange(1).Delete
                                End If
                            End If
                            rg.Collapse wdCollapseStart
                            Set rgTmp = docTmp.BuildingBlockEntries("mpi").Insert(rg)
                            If rgTmp.ShapeRange.Count > 0 Then
                                Set spNew = rgTmp.ShapeRange(1)
                                spNew.LockAspectRatio = msoCTrue
                                spNew.Width = CentimetersToPoints(FOOTER_LOGO_WIDTH)
                                spNew.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                                spNew.Left = CentimetersToPoints(FOOTER_LEFT_TO_PAGE)
                                spNew.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                                spNew.TOP = CentimetersToPoints(FOOTER_TOP_TO_PAGE)
                            End If
                        Case "For", "MPI"           'for those groups, delete MPI logo
                            If rg.ShapeRange.Count > 0 Then
                                Set spNew = rg.ShapeRange(1)
                                If spNew.Title = "MPI" Then
                                    spNew.Delete
                                End If
                            End If
                        Case Else
                        End Select
                        Set rg = doca.Sections(iSec).Footers(iHdr).Range
                        If rg.ShapeRange.Count > 0 Then
                        End If
                    End If
                Next iHdr
            Next iSec
            Exit For
        End If
    Next docTmp
    'for non-bath convert call, write back to normal.dotm
    If sPty = "" Then
        For Each docTmp In oApp.Templates
            If LCase(docTmp) = "normal.dotm" Then
                docTmp.BuiltInDocumentProperties(DOCUMENTPROPERTY) = sBG
                docTmp.Save
                Exit For
            End If
        Next docTmp
    End If
    oApp.ScreenUpdating = True
End Function

'read document property from normal.dotm
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

'read server document property
Public Function ReadServerProperty(doc As Document, pty As String) As String
    ReadServerProperty = ""
    If Trim(pty) = "" Then Exit Function
    On Error Resume Next
    ReadServerProperty = doc.CustomDocumentProperties(pty)
End Function

Public Function FixLogos(sPath As String)
    'check folder exists
    If Trim(sPath) = "" Or Dir(sPath, vbDirectory) = "" Then Exit Function
    Const EXT = "*.docx"    'file type
    Dim File As String
    Dim doc As Document
    'look if folder contains valid file (has a number in filename)
    File = Dir(sPath & EXT, vbNormal)
    While File <> ""
        Set doc = Documents.Open(sPath & File, ConfirmConversions:=False, ReadOnly:=False, Visible:=True)
        DoEvents
        If Not doc Is Nothing Then
            Call SetLogo(doc, SERVERPROPERTY)
            doc.Save
            doc.Close
            Set doc = Nothing
            DoEvents
        End If
        File = Dir  'get next file
    Wend   'end While File<>""
End Function

Sub etset()
    On Error Resume Next
    Debug.Print ActiveDocument.CustomDocumentProperties("Sdfksd234")
    
End Sub
