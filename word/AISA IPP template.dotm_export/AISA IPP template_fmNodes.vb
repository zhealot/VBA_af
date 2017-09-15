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
    If Not ValidateForm Then Exit Sub
    GetNodeByName(GetNodeByName(sCurrent).NextNode).PreviousNode = sCurrent
    sCurrent = GetNodeByName(sCurrent).NextNode
    LoadNode sCurrent
End Sub

Private Sub btnPrevious_Click()
    GetNodeByName(GetNodeByName(sCurrent).PreviousNode).NextNode = sCurrent
    sCurrent = GetNodeByName(sCurrent).PreviousNode
    LoadNode sCurrent
End Sub

Function ValidateForm() As Boolean
'check if controls/fields have been filled
    ValidateForm = False
    If GetNodeByName(sCurrent).NeedAnswer And fmNodes.tbAnswer.Text = "" Then
        MsgBox "Please give an answer to the question."
        fmNodes.tbAnswer.SetFocus
        Exit Function
    End If
    If GetNodeByName(sCurrent).ActionNo > 0 And Not GotAction Then
        MsgBox "Please choose an answer."
        Exit Function
    End If
    GetNodeByName(sCurrent).sAnswer = tbAnswer.Text
    If obYes.Value Then
        GetNodeByName(sCurrent).YesNo = "y"
    ElseIf obNo.Value Then
        GetNodeByName(sCurrent).YesNo = "n"
    Else
        GetNodeByName(sCurrent).YesNo = ""
    End If
    ValidateForm = True
End Function

Private Sub obYes_Click()
    EnableYesNoEvent = True
    YesNo_Click
    EnableYesNoEvent = False
End Sub

Private Sub obNo_Click()
    EnableYesNoEvent = True
    YesNo_Click
    EnableYesNoEvent = False
End Sub

Public Sub YesNo_Click()
    'link following node to current one based on yes/no choice made
    If EnableYesNoEvent Then
        If obYes.Value Then
            GetNodeByName(GetNodeByName(sCurrent).YesNode).PreviousNode = sCurrent
            GetNodeByName(sCurrent).YesNo = "y"
            GetNodeByName(sCurrent).NextNode = GetNodeByName(sCurrent).YesNode
            fmNodes.tbAnswer.Enabled = IIf(GetNodeByName(sCurrent).NeedAnswer, True, False)
            fmNodes.lbTitle.Enabled = IIf(GetNodeByName(sCurrent).NeedAnswer, True, False)
        ElseIf obNo.Value Then
            GetNodeByName(GetNodeByName(sCurrent).NoNode).PreviousNode = sCurrent
            GetNodeByName(sCurrent).YesNo = "n"
            GetNodeByName(sCurrent).NextNode = GetNodeByName(sCurrent).NoNode
            fmNodes.tbAnswer.Enabled = False
            fmNodes.lbTitle.Enabled = False
        Else
            GetNodeByName(sCurrent).YesNo = ""
            GetNodeByName(sCurrent).NextNode = ""
        End If
        sPreNode = GetNodeByName(sCurrent).PreviousNode
    End If
End Sub
Private Sub UserForm_Initialize()
    InitialNodes
    LoadNode aryNodes(0).Name
    sCurrent = aryNodes(0).Name
End Sub
