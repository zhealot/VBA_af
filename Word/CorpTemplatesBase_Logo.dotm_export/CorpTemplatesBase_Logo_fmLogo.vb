VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmLogo 
   Caption         =   "Logo Selection"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   OleObjectBlob   =   "CorpTemplatesBase_Logo_fmLogo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------   -------------------------------------------
' Developed for Ministry for Primary Industries
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             April 2018
' Description:      Set document logo in header/footer according to business group
'-----------------------------------------------------------------------------

Private Sub cbtCancel_Click()
    Unload Me
End Sub

Private Sub cbtOK_Click()
    If sBG = "" Then
        MsgBox "Please choose a Business Group."
        Exit Sub
    End If
    
    Unload Me
    'Dim docA As Document
    Dim rg As Range
    Dim rgTmp As Range
    Dim rgCurrent As Range
    Dim oApp As Word.Application
    Dim tmp As Template
    Set oApp = Word.Application
    oApp.ScreenUpdating = False
    
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
End Sub

Private Sub obBio_Click()
    Call OBClick(obBio)
End Sub


Private Sub obFis_Click()
    Call OBClick(obFis)
End Sub

Private Sub obFor_Click()
    Call OBClick(obFor)
End Sub

Private Sub obMPI_Click()
    Call OBClick(obMPI)
End Sub

Private Sub obNZF_Click()
    Call OBClick(obNZF)
End Sub

Private Sub UserForm_Initialize()
    For Each ctrl In fmLogo.frmImage.Controls
        ctrl.Visible = False
    Next ctrl
    Set docA = ActiveDocument
    Dim tp As Template
    For Each tp In Application.Templates
        If LCase(tp.Name) = "normal.dotm" Then
            Dim pty As Variant
            For Each pty In tp.BuiltInDocumentProperties
                If pty.Name = DOCUMENTPROPERTY Then
                    BusinessGroup = pty
                    Dim ob As control
                    For Each ob In frmBusinessGroup.Controls
                        If Right(ob.Name, Len(ob.Name) - 2) = BusinessGroup Then
                            ob.Value = True
                            Exit For
                        Else
                            ob.Value = False
                        End If
                    Next
                End If
            Next pty
        End If
    Next tp
End Sub
