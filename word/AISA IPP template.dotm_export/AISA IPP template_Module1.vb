Attribute VB_Name = "Module1"
Public aryOptionButtons() As New OptionButtonEvent  'arrry to hold option button object in main frame
Public sSelectedCaption As String       'caption of selected option button in first screen
Public aryNodes() As New oNode
Public aryNodeCnt As Integer
Public sHelpText As String
Public sCurrent As String   'current oNode name
Public sNextNode As String  'next node name
Public sPreNode As String   'previous node name
Public EnableYesNoEvent As Boolean

Function InitialNodes()
    aryNodeCnt = 0
    CreateNode Name:="1", _
                Question:="1 What was the purpose for collecting the information in the first place?", _
                YesNode:="2", _
                NoNode:="2", _
                NeedAnswer:=True, _
                Tip:="Purpose could be redifined to include disclosure for future information holdings (can not be applied retrospectively).", _
                Answer:="", _
                ActionNo:=0, _
                NextNode:="2"
    CreateNode Name:="2", _
                Question:="2 Was this purpose communicated to the individual concerned at the time of collection?", _
                YesNode:="3", _
                NoNode:="exit", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2
    CreateNode Name:="3", _
                Question:="3 I have reasonable grounds to believe the disclosure is a purpose for collecting the information because:", _
                YesNode:="permitted", _
                NoNode:="4", _
                NeedAnswer:=True, _
                Tip:="Remember that an explanation devised in hindsight won't suffice.", _
                Answer:="", _
                ActionNo:=2
    CreateNode Name:="4", _
                Question:="4 I have reasonable grounds for believing that the disclosure is directly related to the purpose for collecting the information because:", _
                YesNode:="permitted", _
                NoNode:="exit", _
                NeedAnswer:=True, _
                Tip:="Remember that an explanation devised in hindsight won't suffice.", _
                Answer:="", _
                ActionNo:=2
            
    '### Permitting and Exit node
    CreateNode Name:="permitted", _
                Question:="Permitting!", _
                YesNode:="Permitting", _
                NoNode:="Permitting", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0
    CreateNode Name:="exit", _
                Question:="Exiting!", _
                YesNode:="exit", _
                NoNode:="exit", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0
End Function

Function CreateDocument(final As Boolean)
'create document based on selections
    Dim doc As Document
    Set doc = ActiveDocument
    
    If final Then
        doc.Content.Delete
        doc.Paragraphs(1).Range.Text = "Your selection: " & sSelectedCaption
    Else
        
    End If
End Function

Function CreateNode(Name As String, Question As String, YesNode As String, _
                    NoNode As String, NeedAnswer As Boolean, Tip As String, _
                    Answer As String, ActionNo As Integer, Optional PreviousNode As String = "", Optional NextNode As String = "", Optional YesNo As String = "") As oNode
'construction oNode object
    Dim nd As New oNode
    With nd
        .Name = Name
        .sQuestion = Question
        .ActionNo = ActionNo
        .sAnswer = Answer
        .sTip = Tip
        .NeedAnswer = NeedAnswer
        .YesNode = YesNode
        .NoNode = NoNode
        .PreviousNode = IIf(PreviousNode = "", "", PreviousNode)
        .YesNo = IIf(YesNo = "", "", YesNo)
        .NextNode = IIf(NextNode = "", "", NextNode)
    End With
    Set CreateNode = nd
    ReDim Preserve aryNodes(aryNodeCnt)
    Set aryNodes(aryNodeCnt) = nd
    aryNodeCnt = aryNodeCnt + 1
End Function


Function LoadNode(nodeName As String)
'populate form using object oNode
    Select Case nodeName
'    Case "permitted"
'        MsgBox "Application permitted."
'    Case "exit"
'        MsgBox "Exit."
    Case Else
        Dim nd As New oNode
        Set nd = GetNodeByName(nodeName)
        fmNodes.lbQuestion.Caption = nd.sQuestion
        fmNodes.tbAnswer.Text = nd.sAnswer
        fmNodes.tbAnswer.Enabled = IIf(nd.NeedAnswer, True, False)  'disable textbox if no text answer needed.
        fmNodes.lbTitle.Enabled = IIf(nd.NeedAnswer, True, False)
        If nd.ActionNo = 0 Then
            fmNodes.fmActions.Enabled = False
            fmNodes.obYes.Enabled = False
            fmNodes.obNo.Enabled = False
            sNextNode = nd.YesNode     'if no choice needed, then link to 'YesNode' by default
        Else
            fmNodes.fmActions.Enabled = True
            fmNodes.obYes.Enabled = True
            fmNodes.obNo.Enabled = True
            Select Case nd.YesNo
            Case "y"
                EnableYesNoEvent = False
                fmNodes.obYes.Value = True
                EnableYesNoEvent = True
            Case "n"
                EnableYesNoEvent = False
                fmNodes.obNo.Value = True
                EnableYesNoEvent = True
            Case Else
                EnableYesNoEvent = False
                fmNodes.obYes.Value = False
                fmNodes.obNo.Value = False
                EnableYesNoEvent = True
            End Select
        End If
        'set button status
        fmNodes.btnPrevious.Enabled = IIf(nd.PreviousNode = "", False, True)
        sHelpText = nd.sTip
        fmNodes.btnHelp.Visible = IIf(sHelpText = "", False, True)
        sCurrent = nodeName
    End Select
End Function

Function GotAction() As Boolean
'check whether Yes/No option button clicked
    If fmNodes.obNo.Value Or fmNodes.obYes.Value Then
        GotAction = True
    Else
        GotAction = False
    End If
End Function

Function GetNodeByName(s As String) As oNode
    Dim nd As New oNode
    Set GetNodeByName = Nothing
    For i = 0 To UBound(aryNodes)
        Set nd = aryNodes(i)
        If nd.Name = s Then
            Set GetNodeByName = nd
            Exit Function
        End If
    Next i
End Function

Function GetNodeIndexByName(s As String) As Integer
    GetNodeIndexByName = -1
    For i = 0 To UBound(aryNodes) - 1
        If aryNodes(i).Name = s Then
            GetNodeIndexByName = i
            Exit For
        End If
    Next i
End Function
