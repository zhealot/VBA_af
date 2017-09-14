VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmNodes 
   Caption         =   "AISA IPP Document - Applying exceptions"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8580
   OleObjectBlob   =   "AISA IPP template_fmNodes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim node1 As oNode
    Set node1 = New oNode
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnHelp_Click()
    MsgBox sHelpText
End Sub

Private Sub btnNext_Click()
    LoadNodeToForm aryNodes(1)
End Sub

Private Sub obPrevious_Click()
    LoadNodeToForm aryNodes(2)
End Sub

Private Sub UserForm_Initialize()
    InitialNodes
    LoadNodeToForm aryNodes(0)
End Sub
