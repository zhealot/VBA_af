VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmBackground 
   Caption         =   "Background Image"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   OleObjectBlob   =   "Cato Bolam_Common_fmBackground.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------   -------------------------------------------
' Developed for Cato Bolam
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             March 2018
' Description:      Check document for error and highlight etc, add cover page & address in footer then export to PDF
'-----------------------------------------------------------------------------


Private Sub cbtCancel_Click()
    Unload Me
End Sub

Private Sub cbtOK_Click()
    If sBackground = "" Then
        MsgBox "Please choose a background image."
        Exit Sub
    End If
    If sBackground = LETTERHEADER And sBackgroundAddress = "" Then
        MsgBox "Please choose an address."
        Exit Sub
    End If
    
    Unload Me
    'Dim docA As Document
    Dim rg As Range
    Dim rgTmp As Range
    Dim rgCurrent As Range
    Dim oApp As Word.Application
    Dim tmp As Template
    Dim sp As Shape
    Set oApp = Word.Application
    
    'set range to whole current page
    Set rg = Selection.Range
    rg.Collapse wdCollapseStart
    If docA.Content.Information(wdNumberOfPagesInDocument) = 1 Then     'in case document has only one page
        Set rg = docA.Content
    Else    'more than one page
        'set rg to the start of current page
        Set rg = rg.GoTo(wdGoToPage, , rg.Information(wdActiveEndPageNumber))
        'if current page is the last page
        If rg.Information(wdActiveEndPageNumber) = docA.Content.Information(wdNumberOfPagesInDocument) Then
            rg.SetRange rg.Start, docA.Content.End
        Else    'not in the last page
            Set rgTmp = rg.GoTo(wdGoToPage, , rg.Information(wdActiveEndPageNumber) + 1)
            rg.SetRange rg.Start, rgTmp.Start - 1
        End If
    End If
    
    'keep current page range
    Set rgCurrent = rg
    
    'check if there's a background image, delete it
    If rg.ShapeRange.Count > 0 Then
        For i = rg.ShapeRange.Count To 1 Step -1
            If rg.ShapeRange(i).Width = rg.PageSetup.PageWidth Then
                rg.ShapeRange(i).Delete 'delete the background image
            End If
        Next i
    End If
     
     'set rg to start of current page
    Set rg = rgCurrent
    rg.Collapse wdCollapseStart
    For Each tmp In oApp.Templates
        On Error Resume Next
        If tmp.Name = ThisDocument.Name Then
            'insert first page cover to background
            Set rg = tmp.BuildingBlockEntries(IIf(sBackground = "AppendixHeader", "LetterHeader", sBackground)).Insert(rg) 'for Appendix Header, use Letter Header image
            'adjust cover page image
            Set sp = rg.ShapeRange(1)
            If Not sp Is Nothing Then
                With sp
                    .WrapFormat.Type = wdWrapBehind
                    .LockAspectRatio = msoFalse
                    .Top = 0
                    .Left = 0
                    .Width = rg.PageSetup.PageWidth
                    If sBackground <> LETTERHEADER And sBackground <> "AppendixHeader" Then
                        .Height = rg.PageSetup.PageHeight
                    End If
                End With
            End If
            
            'insert footer for Letter Header cover
            Set sp = Nothing
            If sBackground = LETTERHEADER Then
                Set rg = rgCurrent
                rg.Collapse wdCollapseStart
                Set rg = tmp.BuildingBlockEntries(LETTERHEADER & sBackgroundAddress).Insert(rg)
                Set sp = rg.ShapeRange(1)
                If Not sp Is Nothing Then
                    With sp
                        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        .WrapFormat.Type = wdWrapBehind
                        .LockAspectRatio = msoFalse
                        .Left = wdShapeLeft
                        .Top = wdShapeBottom
                        .Width = rg.PageSetup.PageWidth
                    End With
                End If
            Else
                'set first page text colour
                Set rg = rgCurrent
                Select Case sCover
                Case "FullCover"
                    rg.Font.ColorIndex = wdWhite
                Case Else
                    rg.Font.ColorIndex = wdBlack
                End Select
            End If  'if sCover = LETTERHEADER
            Exit For
        End If  'tmp.Name = ThisDocument.Name
    Next
End Sub

Private Sub obAppendixHeader_Click()
    Call ChooseCover(obAppendixHeader)
End Sub

Private Sub obFullCover_Click()
    Call ChooseCover(obFullCover)
End Sub

Private Sub obLetterHeader_Click()
    Call ChooseCover(obLetterHeader)
End Sub


Private Sub obReportCoverFormal_Click()
    Call ChooseCover(obReportCoverFormal)
End Sub

Private Sub obReportCoverHeader_Click()
    Call ChooseCover(obReportCoverHeader)
End Sub

Private Sub obWhangarei_Click()
    Call ChooseBackgroundFooter(obWhangarei)
End Sub

Private Sub obHenderson_Click()
    Call ChooseBackgroundFooter(obHenderson)
End Sub

Private Sub obManukau_Click()
    Call ChooseBackgroundFooter(obManukau)
End Sub

Private Sub obOrewa_Click()
    Call ChooseBackgroundFooter(obOrewa)
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Initialize()
    imgFullCover.Visible = False
    imgReportCoverFormal.Visible = False
    imgReportCoverHeader.Visible = False
    imgLetterHeader.Visible = False
    ControlSwitch fmFooter, False
End Sub
