VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Cover & Address"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   OleObjectBlob   =   "Cato Bolam_Common_UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
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
    If sCover = "" Then
        MsgBox "Please choose a cover page."
        Exit Sub
    End If
    If sCover = LETTERHEADER And sAddress = "" Then
        MsgBox "Please choose an address."
        Exit Sub
    End If
    
    Me.Hide
    'Dim docA As Document
    Dim rg As Range
    Dim oApp As Word.Application
    Set oApp = Word.Application
    'look for 1st page range
    Set rg = docA.Content
    If rg.Information(wdNumberOfPagesInDocument) > 1 Then
        If docA.Paragraphs.Count > 1 Then
            'set rg to the start of page 2
            Set rg = rg.GoTo(wdGoToPage, , 2)
            'redefine rg to be of whole page 1
            rg.SetRange 0, rg.Start - 1
        End If
    End If
    'check if there's a background image
    If rg.ShapeRange.Count > 0 Then
        For i = rg.ShapeRange.Count To 1 Step -1
            If rg.ShapeRange(i).Width = rg.PageSetup.PageWidth Then
                rg.ShapeRange(i).Delete 'delete the background image
            End If
        Next i
    End If
    
    'insert cover page
    If sCover <> "NoCover" Then
        'set rg to start of first page
        rg.Collapse wdCollapseStart
        Dim tmp As Template
        For Each tmp In oApp.Templates
            On Error Resume Next
            If tmp.Name = ThisDocument.Name Then
                'insert first page cover to background
                Set rg = tmp.BuildingBlockEntries(sCover).Insert(rg)
                'adjust cover page image
                Dim sp As Shape
                Set sp = rg.ShapeRange(1)
                If Not sp Is Nothing Then
                    With sp
                        .WrapFormat.Type = wdWrapBehind
                        .LockAspectRatio = msoFalse
                        .Top = 0
                        .Left = 0
                        .Width = docA.Sections(1).PageSetup.PageWidth
                        If sCover <> LETTERHEADER Then
                            .Height = docA.Sections(1).PageSetup.PageHeight
                        End If
                    End With
                End If
                'insert singing cover
                If sSigningCover <> "NoSigningCover" And sSigningCover <> "" Then
                    If docA.Content.Information(wdNumberOfPagesInDocument) > 1 Then
                        Set rg = docA.Content
                        Set rg = rg.GoTo(wdGoToPage, , 2)
                        Set rg = tmp.BuildingBlockEntries(sSigningCover).Insert(rg)
                        Set sp = rg.ShapeRange(1)
                        If Not sp Is Nothing Then
                            With sp
                                .WrapFormat.Type = wdWrapBehind
                                .LockAspectRatio = msoFalse
                                .Top = 0
                                .Left = 0
                                .Width = docA.Sections(rg.Information(wdActiveEndSectionNumber)).PageSetup.PageWidth
                                .Height = docA.Sections(rg.Information(wdActiveEndSectionNumber)).PageSetup.PageHeight
                            End With
                        End If
                    End If
                End If
                
                'insert footer for Letter Header cover
                Set sp = Nothing
                If sCover = LETTERHEADER Then
                    Set rg = docA.Content
                    rg.Collapse wdCollapseStart
                    Set rg = tmp.BuildingBlockEntries(LETTERHEADER & sAddress).Insert(rg)
                    Set sp = rg.ShapeRange(1)
                    If Not sp Is Nothing Then
                        With sp
                            .WrapFormat.Type = wdWrapBehind
                            .WrapFormat.DistanceBottom = 0
                            .LockAspectRatio = msoFalse
                            .Left = 0
                            .Width = docA.Sections(1).PageSetup.PageWidth
                        End With
                    End If
                Else
                    'set first page text colour
                    rg.Collapse wdCollapseStart
                    Set rg = rg.GoTo(wdGoToPage, , 2)
                    'set rg to be 1st page
                    rg.SetRange 0, rg.End - 1
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
    End If 'if sCover <> "NoCover"

    'dialogue to save as PDF
    With oApp.Dialogs(wdDialogFileSaveAs)
        .Format = wdFormatPDF
        .Show
    End With
End Sub

Private Sub obFullCover_Click()
    Call ChooseCover(obFullCover)
End Sub

Private Sub obLetterHeader_Click()
    Call ChooseCover(obLetterHeader)
End Sub

Private Sub obNoCover_Click()
    Call ChooseCover(obNoCover)
End Sub

Private Sub obNoSigningCover_Click()
    Call ChooseSigningCover(obNoSigningCover)
End Sub

Private Sub obReportCoverFormal_Click()
    Call ChooseCover(obReportCoverFormal)
End Sub

Private Sub obReportCoverHeader_Click()
    Call ChooseCover(obReportCoverHeader)
End Sub

Private Sub obSigningCover1_Click()
    Call ChooseSigningCover(obSigningCover1)
End Sub

Private Sub obSigningCover2_Click()
    Call ChooseSigningCover(obSigningCover2)
End Sub

Private Sub obWhangarei_Click()
    Call ChooseFooter(obWhangarei)
End Sub

Private Sub obHenderson_Click()
    Call ChooseFooter(obHenderson)
End Sub

Private Sub obManukau_Click()
    Call ChooseFooter(obManukau)
End Sub

Private Sub obOrewa_Click()
    Call ChooseFooter(obOrewa)
End Sub

Private Sub UserForm_Initialize()
    imgFullCover.Visible = False
    imgReportCoverFormal.Visible = False
    imgReportCoverHeader.Visible = False
    imgLetterHeader.Visible = False
    ControlSwitch fmFooter, False
    
    'check if document has more than 1 page to enable 'Signing page' function
    If ActiveDocument.Content.Information(wdNumberOfPagesInDocument) < 2 Then
        'fmSigningPage.Caption = "Document has only one page"
        ControlSwitch fmSigningPage, False
        sSigningCover = ""
    Else
        'fmSigningPage.Caption = "Signing Page"
        ControlSwitch fmSigningPage, True
    End If

End Sub
