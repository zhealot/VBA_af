VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmLogo 
   Caption         =   "Logo Selection"
   ClientHeight    =   5170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9870
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
    Call SetLogo(ActiveDocument)
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
    'read Business Group name from normal.dotm
    sBG = ReadBG
    'set option button
    Dim ob As control
    For Each ob In frmBusinessGroup.Controls
        If Right(ob.Name, Len(ob.Name) - 2) = sBG Then
            ob.Value = True
            Exit For
        Else
            ob.Value = False
        End If
    Next
End Sub
