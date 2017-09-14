Attribute VB_Name = "Module1"
Public aryOptionButtons() As New OptionButtonEvent  'arrry to hold option button object in main frame
Public sSelectedCaption As String       'caption of selected option button in first screen
Public aryNodes() As New oNode
Public aryNodeCnt As Integer
Public sHelpText As String

Function InitialNodes()
    aryNodeCnt = 0
    Dim ndNew As New oNode
    Set ndNew = CreateNode(Name:="001", _
                           Question:="What was the purpose for collecting the information in the first place?", _
                           Yes:="002", _
                           No:="", _
                           NeedAnswer:=True, _
                           Tip:="Purpose could be redifined to include disclosure for future information holdings (can not be applied retrospectively).", _
                           Answer:="", _
                           ActionNo:=0, _
                           Previous:=Nothing, _
                           Nxt:=Nothing)
    Set ndNew = CreateNode(Name:="002", _
                           Question:="Was this purpose communicated to the individual concerned at the time of collection?", _
                           Yes:="003", _
                           No:="exit", _
                           NeedAnswer:=False, _
                           Tip:="", _
                           Answer:="", _
                           ActionNo:=2, _
                           Previous:=Nothing, _
                           Nxt:=Nothing)
    Set ndNew = CreateNode(Name:="003", _
                           Question:="I have reasonable grounds to believe the disclosure is a purpose for collecting the information because:", _
                           Yes:="permitted", _
                           No:="004", _
                           NeedAnswer:=True, _
                           Tip:="Remember that an explanation devised in hindsight won't suffice.", _
                           Answer:="", _
                           ActionNo:=2, _
                           Previous:=Nothing, _
                           Nxt:=Nothing)
    Set ndNew = CreateNode(Name:="004", _
                           Question:="I have reasonable grounds for believing that the disclosure is directly related to the purpose for collecting the information because:", _
                           Yes:="permitted", _
                           No:="005", _
                           NeedAnswer:=True, _
                           Tip:="Remember that an explanation devised in hindsight won't suffice.", _
                           Answer:="", _
                           ActionNo:=2, _
                           Previous:=Nothing, _
                           Nxt:=Nothing)
    Debug.Print UBound(aryNodes)
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

Function CreateNode(Name As String, Question As String, Yes As String, No As String, NeedAnswer As Boolean, Tip As String, Answer As String, ActionNo As Integer, Previous As oNode, Nxt As oNode) As oNode
    Dim nd As New oNode
    With nd
        .ActionNo = ActionNo
        .Name = Name
        Set .oNext = Nxt
        Set .oPre = Previous
        .sAnswer = Answer
        .sQuestion = Question
        .sTip = Tip
        .NeedAnswer = NeedAnswer
    End With
    Set CreateNode = nd
    ReDim Preserve aryNodes(aryNodeCnt)
    Set aryNodes(aryNodeCnt) = nd
    aryNodeCnt = aryNodeCnt + 1
End Function


Function LoadNodeToForm(nd As oNode)
'populate form using object oNode
    fmNodes.lbQuestion.Caption = nd.sQuestion
    fmNodes.tbAnswer.Text = nd.sAnswer
    If nd.ActionNo = 0 Then
        fmNodes.fmActions.Enabled = False
        fmNodes.obYes.Enabled = False
        fmNodes.obNo.Enabled = False
    Else
        fmNodes.fmActions.Enabled = True
        fmNodes.obYes.Enabled = True
        fmNodes.obNo.Enabled = True
    End If
    sHelpText = nd.sTip
End Function
