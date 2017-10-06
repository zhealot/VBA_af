VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "AISA IPP Document - Sharing Vehicles"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   OleObjectBlob   =   "AISA IPP template_frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim ob As MSForms.OptionButton
    sSelectedCaption = ""
    For i = 0 To Frame1.Controls.Count
        On Error Resume Next
        If InStr(TypeName(Frame1.Controls(i)), "OptionButton") > 0 Then
            Set ob = Frame1.Controls(i)
            If ob.Value Then
                sSelectedCaption = ob.Caption
                Exit For
            End If
        End If
    Next i
    If sSelectedCaption = "" Then
        MsgBox "Please choose an option."
        Exit Sub
    End If
    Me.Hide
    If CommandButton1.Caption = "Yes" Then
        Call CreateDocument("1")
    Else
        fmNodes.Show
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim cnt As Integer
    For Each ctr In Frame1.Controls
        If TypeName(ctr) = "OptionButton" Then
            ReDim Preserve aryOptionButtons(cnt)
            Set aryOptionButtons(cnt).oOBEvents = ctr
            cnt = cnt + 1
        End If
    Next
End Sub
