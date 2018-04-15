Attribute VB_Name = "Allfields"
Option Explicit
Dim oWithEvent As New EventsClass
'module to hold ribbon button code
Const Style1 = "Instructional Text"
Const Style2 = "Instructional Text Bullets"
Const SERVERPROPERTY = "MPIPortfolio"  'property to read from server
Const HEADER_TOP_TO_PAGE = 1.2
Const HEADER_RIGHT_TO_PAGE = 1.1  'distance of header logo right side  to page right side, in centimeter
Const FOOTER_BOTTOM_TO_PAGE = 1.13  'distance of footer logo bottom side to page bottom side, in centimeter
Const FOOTER_LEFT_TO_PAGE = 1.2
Const HEADER_LOGO_WIDTH = 8.6
Const FOOTER_LOGO_WIDTH = 6.2
Public Const DOCUMENTPROPERTY = "Category" 'document property name to hold business group
Public sBG As String 'short business group name
Public docA As Document

Public Sub AutoExec()
    Set oWithEvent.oWdApp = Word.Application
End Sub

Public Function SetLogo(docA As Document, Optional sPty As String = "")
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
    If sBG = "" Then
        sBG = ReadBG
    End If
    If sPty <> "" Then
        sBG = ReadServerProperty(docA, SERVERPROPERTY)
    End If
    If sBG = "" Then
        Exit Function
    End If
    'set range to whole current page
    Set rg = docA.Content
    rg.Collapse wdCollapseStart
    Dim SeekView As Long
    Dim docTmp As Template
    Dim spNew As Shape  'new shape
    For Each docTmp In oApp.Templates
        If docTmp.Name = ThisDocument.Name Then
            'set first page header if any, otherwise set section 1 primary header
            If docA.Sections(1).Headers(wdHeaderFooterFirstPage).Exists Then
                Set rg = docA.Sections(1).Headers(wdHeaderFooterFirstPage).Range
            Else
                Set rg = docA.Sections(1).Headers(wdHeaderFooterPrimary).Range
            End If
            'delete existing logo
            If rg.ShapeRange.Count > 0 Then
                rg.ShapeRange(1).Delete
            End If
            rg.Collapse wdCollapseEnd
            Set rgTmp = docTmp.BuildingBlockEntries(sBG).Insert(rg)
            Set spNew = rgTmp.ShapeRange(1)
            spNew.LockAspectRatio = msoTrue
            spNew.Width = CentimetersToPoints(HEADER_LOGO_WIDTH)
            spNew.WrapFormat.Type = wdWrapSquare
            spNew.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            spNew.Left = docA.Sections(1).PageSetup.PageWidth - spNew.Width - CentimetersToPoints(HEADER_RIGHT_TO_PAGE)
            spNew.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            spNew.Top = CentimetersToPoints(HEADER_TOP_TO_PAGE)
            
            'deal with the footer part
            'delete all images in footer in all sections
            Dim sec As Section
            For Each sec In docA.Sections
                Dim iFt As Integer
                For iFt = 1 To 3
                    If sec.Footers(iFt).Exists Then
                        Set rg = sec.Footers(iFt).Range
                        If rg.ShapeRange.Count > 0 Then
                            Dim i As Integer
                            For i = rg.ShapeRange.Count To 1 Step -1
                                'If rg.ShapeRange(i).Title = "MPI" Then
                                    rg.ShapeRange(i).Delete
                                'End If
                            Next i
                        End If
                    End If
                Next iFt
            Next sec
            'set range
            If docA.Sections(1).Footers(wdHeaderFooterFirstPage).Exists Then
                Set rg = docA.Sections(1).Footers(wdHeaderFooterFirstPage).Range
            Else
                Set rg = docA.Sections(1).Footers(wdHeaderFooterPrimary).Range
            End If
            'insert footer logo is necessary
            Select Case sBG
            Case "Bio", "Fis", "NZF"    'for those groups, insert MPI logo
                rg.Collapse wdCollapseStart
                Set rgTmp = docTmp.BuildingBlockEntries("mpi").Insert(rg)
                If rgTmp.ShapeRange.Count > 0 Then
                    Set spNew = rgTmp.ShapeRange(1)
                    spNew.LockAspectRatio = msoCTrue
                    spNew.Width = CentimetersToPoints(FOOTER_LOGO_WIDTH)
                    spNew.WrapFormat.Type = wdWrapSquare
                    spNew.WrapFormat.AllowOverlap = False
                    spNew.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    spNew.Left = CentimetersToPoints(FOOTER_LEFT_TO_PAGE)
                    spNew.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    spNew.Top = docA.Sections(1).PageSetup.PageHeight - spNew.Height - CentimetersToPoints(FOOTER_BOTTOM_TO_PAGE)
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
    Dim s As String
    s = doc.CustomDocumentProperties(pty)
    If InStr(s, "Biosecurity") > 0 Then
        ReadServerProperty = "Bio"
    ElseIf InStr(s, "Fisheries") > 0 Then
        ReadServerProperty = "Fis"
    ElseIf InStr(s, "Forestry") > 0 Then
        ReadServerProperty = "For"
    ElseIf InStr(s, "Food Safety") > 0 Then
        ReadServerProperty = "NZF"
    ElseIf InStr(s, "Primary Industries") > 0 Then
        ReadServerProperty = "MPI"
    End If
End Function

Public Function FixLogos(sPath As String)
    Const EXT = "*.docx"    'file type
    Dim File As String
    Dim doc As Document
    
    If LCase(Left(sPath, 4)) = "http" Then  'for Sharepoint server address
        Dim oNet As Object
        Dim fs As Object
        Dim oFolder As Object
        Dim oFile As Object
        Set oNet = CreateObject("WScript.Network")
        Set fs = CreateObject("Scripting.FileSystemObject")
        oNet.mapnetworkdrive "A:", sPath
        Set oFolder = fs.GetFolder("A:")

        For Each oFile In oFolder.Files
            Set doc = Documents.Open(sPath & oFile, ConfirmConversions:=False, ReadOnly:=False, Visible:=True)
            DoEvents
            If Not doc Is Nothing Then
                Call SetLogo(doc, SERVERPROPERTY)
                doc.Save
                doc.Close
                Set doc = Nothing
                DoEvents
            End If
        Next oFile
        oNet.removenetworkdrive "A:"
    Else                                    'for local folder address
        'check folder exists
        If Trim(sPath) = "" Or Dir(sPath, vbDirectory) = "" Then Exit Function
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
    End If
End Function

Sub etset()
    On Error Resume Next
    Debug.Print ActiveDocument.CustomDocumentProperties("Sdfksd234")
End Sub
