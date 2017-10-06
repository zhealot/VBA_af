Attribute VB_Name = "Module1"
'-----------------------------------------------------------------------------
' Developed for DIA
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             October 2017
' Description:      Implement decision making tree for IPP
'-----------------------------------------------------------------------------

Public aryOptionButtons() As New OptionButtonEvent  'arrry to hold option button object in main frame
Public sSelectedCaption As String       'caption of selected option button in first screen
Public aryNodes() As New oNode
Public aryNodeCnt As Integer
Public sHelpText As String
Public sCurrent As String   'current oNode name
Public sNextNode As String  'next node name
Public sPreNode As String   'previous node name
Public EnableYesNoEvent As Boolean
Public Const DefaultAnswerText = "Give your answer later."
Public Const PlaceHolderText = "Space to write more"
Public Const QuestionStyle = "Question" 'Word style name for question
Public Const AnswerStyle = "Answer"     'Word style name for yes/no answer
Public Const FirstNode = "51"            'name of node to start with
Public Const PreFixString = "Your choice(s) indicate that:  Disclosure is permitted ("  'prefix wording for IPP exception clauses

Function InitialNodes()
    ReDim aryNodes(0)
    aryNodeCnt = 0
    '### Permitting and Exit node
    CreateNode Name:="permitted", _
                Question:="Your application is permitted.", _
                YesNode:="", _
                NoNode:="", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0
    CreateNode Name:="exit", _
                Question:="Your application is not permitted!", _
                YesNode:="", _
                NoNode:="", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0
    
    'Normal nodes
    CreateNode Name:="1", _
                Question:="What was the purpose for collecting the information in the first place?", _
                YesNode:="2", _
                NoNode:="2", _
                NeedAnswer:=True, _
                Tip:="Purpose could be redifined to include disclosure for future information holdings (can not be applied retrospectively).", _
                Answer:="", _
                ActionNo:=0, _
                NextNode:="2"
                
    CreateNode Name:="2", _
                Question:="Was this purpose communicated to the individual concerned at the time of collection?", _
                YesNode:="3", _
                NoNode:="5", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="For example:" & vbNewLine & Chr(149) & "There is a statement on our arrival form that explains this type of disclosure, or" & vbNewLine & Chr(149) & "We have a legal opinion that this type of disclosure  is implicit in the purpose for collection", _
                ActionNo:=2
                
    CreateNode Name:="3", _
                Question:="I have reasonable grounds to believe the disclosure is a purpose for collecting the information because:", _
                YesNode:="54", _
                NoNode:="4", _
                NeedAnswer:=True, _
                Tip:="Remember that an explanation devised in hindsight won't suffice." & "Whether or not a purpose included disclosure, or whether a disclosure is directly related to the purposeis a question of fact." & vbNewLine & "(Director of Human Rights Proceedings v Crampton [2015] NZHRRT 35 at [81-82])" & vbNewLine & "That makes it advisable to document the purpose for collecting, obtaining, or creating information, and to note the reasons for disclosing it.", _
                Answer:="", _
                ActionNo:=2, _
                YesText:=PreFixString & "IPP 11(a))"
                
    CreateNode Name:="4", _
                Question:="I have reasonable grounds for believing that the disclosure is directly related to the purpose for collecting the information because:", _
                YesNode:="54", _
                NoNode:="5", _
                NeedAnswer:=True, _
                Tip:="Whether or not a purpose included disclosure, or whether a disclosure is directly related to the purposeis a question of fact." & vbNewLine & "(Director of Human Rights Proceedings v Crampton [2015] NZHRRT 35 at [81-82])" & vbNewLine & "That makes it advisable to document the purpose for collecting, obtaining, or creating information, and to note the reasons for disclosing it.", _
                Answer:="", _
                ActionNo:=2, _
                YesText:=PreFixString & "IPP 11(a))"
                
    CreateNode Name:="5", _
                Question:="Is the disclosure to the individual concerned?", _
                YesNode:="54", _
                NoNode:="6", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesTextBox:=True, _
                YesText:=PreFixString & "IPP 11(c))"
                
    CreateNode Name:="6", _
                Question:="Is the disclosure authorised by the individual concerned?", _
                YesNode:="54", _
                NoNode:="7", _
                NeedAnswer:=False, _
                Tip:="How recent is the authorisation?" & vbNewLine & "Should a new authorisation be sought?", _
                Answer:="", _
                ActionNo:=2, _
                YesTextBox:=True, _
                YesText:=PreFixString & "IPP 11(d))"
                
    CreateNode Name:="7", _
                Question:="Does the information come from a publicly available publication?", _
                YesNode:="54", _
                NoNode:="8", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesTextBox:=True, _
                YesText:=PreFixString & "IPP 11(b))"

    CreateNode Name:="8", _
                Question:="Is it going to be used in a way that will indentify the individual?", _
                YesNode:="9", _
                NoNode:="54", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                NoTextBox:=True, _
                NoText:=PreFixString & "IPP 11(h)(i))"
    
    '### this node needs a 'No' branch
    CreateNode Name:="9", _
                Question:="Is it going to be used for statistical or research purpose?", _
                YesNode:="10", _
                NoNode:="54", _
                NeedAnswer:=False, _
                Tip:="Information does not have to be de-identified at point of disclosure, as long as the published research doesn't identify individuals.", _
                Answer:="", _
                ActionNo:=2, _
                NoTextBox:=True, _
                NoText:=PreFixString & "IPP 11())"

    CreateNode Name:="10", _
                Question:="Will the published research identify individuals?", _
                YesNode:="11", _
                NoNode:="54", _
                NeedAnswer:=False, _
                Tip:="May need something here about what identification means...", _
                Answer:="", _
                ActionNo:=2, _
                NoTextBox:=True, _
                NoText:=PreFixString & "IPP 11(h)(ii))"

    CreateNode Name:="11", _
                Question:="Do the individual consent to the disclosure?", _
                YesNode:="54", _
                NoNode:="12", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesTextBox:=True, _
                YesText:=PreFixString & "IPP 11(a))"

    CreateNode Name:="12", _
                Question:="Is disclosure part of the sale or disposition of a business as a going concern?", _
                YesNode:="54", _
                NoNode:="13", _
                NeedAnswer:=False, _
                Tip:="E.g. the sale of a retail business or a professional firm (e.g. law firm, accountancy) includes its customer list." & vbNewLine & " This exception DOES NOT permit:" & vbNewLine & Chr(149) & " Sale of a customer list without the business also being sold." & vbNewLine & Chr(149) & "sale of a customer list to defray debts in a receivership or liquidation.", _
                Answer:="", _
                ActionNo:=2, _
                YesTextBox:=True, _
                YesText:=PreFixString & "IPP 11(g))"

    CreateNode Name:="13", _
                Question:="Has the Privacy Commissioner authorised me the disclosure the information?", _
                YesNode:="54", _
                NoNode:="14", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesTextBox:=True, _
                YesText:=PreFixString & "IPP 11(i))"

    CreateNode Name:="14", _
                Question:="Is disclosure necessary to avoid prejudice to maintenance of the law?", _
                YesNode:="17", _
                NoNode:="17", _
                NeedAnswer:=False, _
                Tip:="Maintenance of the law includes:" & vbNewLine & Chr(149) & " Prevention - e.g." & vbNewLine & Chr(149) & " Detection - e.g. checking with an agency to see whether an employee has wrongfully accessed an information systems, to verify allegations made (Tan v NZ Police [2016] NZHRRT 32)" & vbNewLine & Chr(149) & " Investigation - e.g." & vbNewLine & Chr(149) & " Prosecution - e.g." & vbNewLine & Chr(149) & " Punishment - e.g. disclosure details of fines to Ministry of Justice collections unit for enforcement.", _
                Answer:="", _
                ActionNo:=0
                
    CreateNode Name:="17", _
                Question:="The Law in issue is:", _
                YesNode:="18", _
                NoNode:="18", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="18", _
                Question:="The law is enforced by:", _
                YesNode:="19", _
                NoNode:="19", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="19", _
                Question:="This agency is a public sector agency.", _
                YesNode:="15", _
                NoNode:="20", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2

    CreateNode Name:="15", _
                Question:="I believe the disclosure is necessary because:", _
                YesNode:="16", _
                NoNode:="16", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="16", _
                Question:="I have reasonable grounds for my belief because: ", _
                YesNode:="54", _
                NoNode:="20", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesText:=PreFixString & "IPP 11(e)(i))"
                
    CreateNode Name:="20", _
                Question:="Is disclosure necessary for enforcement of a law imposing a percuniary penalty?", _
                YesNode:="21", _
                NoNode:="23", _
                NeedAnswer:=False, _
                Tip:="Pecuniary penalties are monetary penalties imposed by statue. They are intended to punish and deter contravention of the law. They may be issued in civil or criminal proceedings.", _
                Answer:="", _
                ActionNo:=2

    CreateNode Name:="21", _
                Question:="I believe the disclosure is necessary because:  ", _
                YesNode:="22", _
                NoNode:="22", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="22", _
                Question:="I have reasonable grounds for my belief because: ", _
                YesNode:="54", _
                NoNode:="24", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesText:=PreFixString & "IPP 11(e)(ii))"

    CreateNode Name:="23", _
                Question:="There is no such law", _
                YesNode:="24", _
                NoNode:="24", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="24", _
                Question:="Is disclosure necessary for the protection of the public revenue?", _
                YesNode:="25", _
                NoNode:="26", _
                NeedAnswer:=False, _
                Tip:="E.g. disclosure is to assess tax liabilities identify benefit fraud, enforce child support payments or payment of infringements or court fines.", _
                Answer:="", _
                ActionNo:=2

    CreateNode Name:="25", _
                Question:="The public revenue in issue is:", _
                YesNode:="27", _
                NoNode:="27", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="27", _
                Question:="I believe the disclosure is necessary because:", _
                YesNode:="28", _
                NoNode:="28", _
                NeedAnswer:=True, _
                Tip:="  ", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="28", _
                Question:="I have reasonable grounds for my belief because:", _
                YesNode:="54", _
                NoNode:="29", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesText:=PreFixString & "IPP 11(e)(iii))"

    CreateNode Name:="26", _
                Question:="The public revenue is not in issue:", _
                YesNode:="29", _
                NoNode:="29", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="29", _
                Question:="Is the disclosure necessary for the conduct of proceedings before any court or tribunal?", _
                YesNode:="30", _
                NoNode:="34", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2

    CreateNode Name:="30", _
                Question:="The proceedings have started or are reasonable in contemplation", _
                YesNode:="31", _
                NoNode:="34", _
                NeedAnswer:=False, _
                Tip:="Reasonable in contemplation means[to come]...", _
                Answer:="", _
                ActionNo:=2

    CreateNode Name:="31", _
                Question:="The proceedings are:", _
                YesNode:="32", _
                NoNode:="32", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="32", _
                Question:="I believe the disclosure is necessary because:", _
                YesNode:="33", _
                NoNode:="33", _
                NeedAnswer:=True, _
                Tip:="Necessary means 'needed or required' in the circumstances, not just 'desirable or expedient'." & vbNewLine & "However, 'needed or required' is something less than 'indispensible or essential'.", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="33", _
                Question:="I have reasonable grounds for my belief because:", _
                YesNode:="54", _
                NoNode:="34", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesText:=PreFixString & "IPP 11(e)(iv))"

    '###need a no branch
    CreateNode Name:="34", _
                Question:="Is disclosure necessary to prevent or lessen a serious threat?", _
                YesNode:="35.1", _
                NoNode:="35.1", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="35.1", _
                Question:="Is it a threat to public health or public safety?", _
                YesNode:="36", _
                NoNode:="35.2", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2

    CreateNode Name:="35.2", _
                Question:="Is it a threat to a specific person?", _
                YesNode:="36", _
                NoNode:="exit", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                NoText:="Return to the beginning to see if other exceptions apply. If not, authorisation will be needed for the disclosure."

    CreateNode Name:="36", _
                Question:="Is it a serious threat?", _
                YesNode:="37", _
                NoNode:="exit", _
                NeedAnswer:=False, _
                Tip:="A 'serious' threat is one that the agency reasonably believes is serious based on three factors:", _
                Answer:="", _
                ActionNo:=2, _
                NoText:="Return to the beginning to see if other exceptions apply. If not, authorisation will be needed for the disclosure."

    CreateNode Name:="37", _
                Question:="How likely is it that the threat will come to pass?", _
                YesNode:="38", _
                NoNode:="38", _
                NeedAnswer:=True, _
                Tip:="Is the threat very likely to occur? Is it possible, or is there only a remote chance?", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="38", _
                Question:="How serious will the consequences be if the threat comes to pass?", _
                YesNode:="39", _
                NoNode:="39", _
                NeedAnswer:=True, _
                Tip:="Will the consequences be felt by one person or many?" & vbNewLine & "Is anyone likely to die or be injured as a result of the threat?" & vbNewLine & "Are people likely to have their identities, financial details, or money stolen as a result of the threat?", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="39", _
                Question:="When is the threat likely to come to pass?", _
                YesNode:="40", _
                NoNode:="40", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=0

    CreateNode Name:="40", _
                Question:="I have reasonable grounds for my assessment that the threat is serious because: ", _
                YesNode:="41", _
                NoNode:="exit", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                NoText:="Return to the beginning to see if other exceptions apply. If not, authorisation will be needed for the disclosure."

    CreateNode Name:="41", _
                Question:="I believe the disclosure is necessary because: ", _
                YesNode:="42", _
                NoNode:="42", _
                NeedAnswer:=True, _
                Tip:="Necessary means 'needed or required' in the circumstances, not just 'desirable or expedient'." & vbNewLine & "However, 'needed or required' is something less than 'indispensible or essential'.", _
                Answer:="", _
                ActionNo:=0
                
    CreateNode Name:="42", _
                Question:="I have reasonable grounds for my belief because: ", _
                YesNode:="54", _
                NoNode:="exit", _
                NeedAnswer:=True, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesText:=PreFixString & "IPP 11(f))", _
                NoText:="Return to the beginning to see if other exceptions apply. If not, authorisation will be needed for the disclosure (e.g. AISA, Schedule 4A or 5 entry, bespoke legislation."
    
    CreateNode Name:="51", _
                Question:="Does legislation other than the Privacy Act prevent or regulate disclosure?", _
                YesNode:="52", _
                NoNode:="53", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2
    
    CreateNode Name:="52", _
                Question:="Is information subject to Tax Administration Act, Senior Courts Act, District Courts Act, or Births, Deaths, Marriages, and Relationships Registration Act?", _
                YesNode:="exit", _
                NoNode:="exit", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                YesText:="AISA required", _
                NoText:="Comply with the law (or seek an amendment)"
                
    CreateNode Name:="53", _
                Question:="Do any of the IPP exceptions clearly apply?", _
                YesNode:="1", _
                NoNode:="exit", _
                NeedAnswer:=False, _
                Tip:="IPP 11 sets out a general rule that information should not be disclosed unless disclosure is one of the purposes for having the information in the first place." & vbNewLine & "IPP 11 sets out a range of exceptions to this general rule. In proceedings under the Privacy Act, defendants have to prove an exception applies so it makes sense to keep a record of the decision and its rationale.", _
                Answer:="", _
                ActionNo:=2, _
                NoText:="Authorisation needed (e.g. AISA, Schedule 4A or 5 entry, Privacy Act Code of Practice, bespoke legislation, s 54 authorisation)"
    
    CreateNode Name:="54", _
                Question:="What are the threshold criteria? For bulk or automated releases, can the threshold criteria be applied automatically?", _
                YesNode:="55", _
                NoNode:="permitted", _
                NeedAnswer:=False, _
                Tip:="Consider using a memorandum of understanding to: specify how information sharing will work;" & vbNewLine & "identify how accuracy of disclosed information will be ensured; specify how disclosed information will be used. Consider publishing memorandums of understanding to enhance transparency.", _
                Answer:="", _
                ActionNo:=2, _
                NoText:="Case by case disclosure will be required."
                
    CreateNode Name:="55", _
                Question:="Can the information at issue be served from other information, so disclosure is limited to the relevant information?", _
                YesNode:="56", _
                NoNode:="exit", _
                NeedAnswer:=False, _
                Tip:="", _
                Answer:="", _
                ActionNo:=2, _
                NoText:="Authorisation needed (e.g. AISA, Schedule 4A or 5 entry, Privacy Act Code of Practice, bespoke legislation, s 54 authorisation)"
    
    '###
    CreateNode Name:="56", _
                Question:="Is there a risk that the information disclosure could cause harm to an individual?", _
                YesNode:="exit", _
                NoNode:="permitted", _
                NeedAnswer:=False, _
                Tip:="Harm includes taking adverse action against a person (e.g. stopping a benefit, imposing a sanction). Adverse action is justifiable if based on accurate information, and accompanied by natural justice. Harm includes significant distress or humiliation, material losses, damage to reputation. Harm is particularly likely if information is inaccurate.", _
                Answer:="", _
                ActionNo:=2, _
                YesText:="Reduce risk by building in natural justice (for adverse action) and safeguards to improve accuracy", _
                NoText:="Disclosure in accordance with IPP exceptions"
                
End Function

Function CreateDocument(stage As String)
'create document based on selections
    Dim doc As Document
    Set doc = ActiveDocument
    fmNodes.Hide
    Dim bm As Bookmark
    Dim rg As Range
    Set rg = doc.Paragraphs.Last.Range
    If stage = "1" Then
    'finish at 1st pop up
        rg.Text = sSelectedCaption & vbNewLine & vbTab & "Yes." & vbNewLine & "Your application is permitted."
    Else
    'rest decision tree, spit out all questions and answer/choices
        Dim nd As oNode
        Set nd = GetNodeByName(FirstNode)
        Do While nd.Name <> "exit" And nd.Name <> "permitted"
            doc.Paragraphs.Add
            doc.Paragraphs.Last.Range.Text = nd.sQuestion '& vbNewLine & vbTab & IIf(nd.ActionNo > 0, IIf(nd.YesNo = "y", "Yes: ", "No."), "") & nd.sAnswer
            doc.Paragraphs.Last.Range.Style = QuestionStyle
            If nd.ActionNo > 0 Then
                doc.Paragraphs.Add
                doc.Paragraphs.Last.Range.Style = AnswerStyle
                doc.Paragraphs.Last.Range.Text = IIf(nd.ActionNo > 0, IIf(nd.YesNo = "y", "Yes: ", "No."), "")
            End If
            'for those needs a statement before next question
            If (nd.YesNo = "y" And nd.YesText <> "") Or (nd.YesNo = "n" And nd.NoText <> "") Then
                doc.Paragraphs.Add
                doc.Paragraphs.Last.Style = AnswerStyle
                doc.Paragraphs.Last.Range.Text = IIf(nd.YesNo = "y", nd.YesText, nd.NoText)
            End If
            If nd.NeedAnswer Or (nd.YesNo = "y" And nd.YesTextBox) Or (nd.YesNo = "n" And nd.NoTextBox) Then
                doc.Paragraphs.Add
                doc.Paragraphs.Last.Range.Style = AnswerStyle
                Set rg = doc.AttachedTemplate.BuildingBlockEntries("IPP_AnswerBox_Blank").Insert(doc.Paragraphs.Last.Range, True)
                If nd.sAnswer = "" Then
                    If rg.Tables.Count > 0 Then
                        rg.Tables(1).Cell(1, 1).Range.Text = PlaceHolderText
                    End If
                ElseIf nd.sAnswer <> DefaultAnswerText And Left(nd.sAnswer, 3) <> "IPP" Then
                    If rg.Tables.Count > 0 Then
                        rg.Tables(1).Cell(1, 1).Range.Text = nd.sAnswer
                    End If
                End If
                rg.Editors.Add wdEditorEveryone
            End If
            Set nd = GetNodeByName(nd.NextNode)
        Loop
        doc.Paragraphs.Add
        Set rg = doc.Content
        rg.Collapse wdCollapseEnd
        If nd.Name = "exit" Then
            rg.Text = "Your application is not permitted."
        Else
            rg.Text = "Your application is permitted."
        End If
    End If
    doc.Protect wdAllowOnlyReading
End Function

Function CreateNode(Name As String, Question As String, YesNode As String, _
                    NoNode As String, NeedAnswer As Boolean, Tip As String, _
                    Answer As String, ActionNo As Integer, Optional PreviousNode As String = "", _
                    Optional NextNode As String = "", Optional YesNo As String = "", _
                    Optional YesTextBox As Boolean = False, Optional NoTextBox As Boolean = False, _
                    Optional YesText As String = "", Optional NoText As String = "") As oNode
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
        .YesTextBox = YesTextBox
        .NoTextBox = NoTextBox
        .YesText = YesText
        .NoText = NoText
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
        fmNodes.lbQuestion.Caption = nd.Name & " " & nd.sQuestion '###put node name before question text
        fmNodes.lbAnswer.Caption = IIf(nd.NeedAnswer, DefaultAnswerText, "") 'nd.sAnswer
        fmNodes.lbAnswer.Enabled = True 'IIf(nd.NeedAnswer, True, False)  'disable textbox if no text answer needed.
        'fmNodes.lbTitle.Enabled = False 'IIf(nd.NeedAnswer, True, False)
        If nd.ActionNo = 0 Then
            fmNodes.fmActions.Enabled = False
            fmNodes.obYes.Enabled = False
            fmNodes.obNo.Enabled = False
            sNextNode = nd.YesNode     'if no choice needed, then link to 'YesNode' by default
            nd.NextNode = nd.YesNode
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
        'set text of permitted / exit node
'        If nd.YesText <> "" Then
'            GetNodeByName(nd.YesNode).sAnswer = nd.YesText
'        End If
'        If nd.NoText <> "" Then
'            GetNodeByName(nd.NoNode).sAnswer = nd.NoText
'        End If
        
'        If nd.YesNode = "permitted" Or nd.NoNode = "permitted" Then
'            GetNodeByName("permitted").sAnswer = nd.YesText
'        End If
'        If nd.YesNode = "exit" Or nd.NoNode = "exit" Then
'        End If

        'set text in answer text box
        If nd.sAnswer <> "" And Left(nd.sAnswer, 3) <> "IPP" Then
            fmNodes.lbAnswer.Caption = DefaultAnswerText 'nd.sAnswer
        End If
    End Select
    '###set button capiton
    If nodeName = "exit" Or nodeName = "permitted" Then
        fmNodes.btnNext.Caption = "Finish"
        fmNodes.lbAnswer.Enabled = False
        fmNodes.fmActions.Enabled = False
        fmNodes.obYes.Enabled = False
        fmNodes.obNo.Enabled = False
    Else
        fmNodes.btnNext.Caption = "Next"
    End If
    'set pop up form title
    If IsNumeric(nodeName) Then
        If 0 < nodeName < 14 Then
            fmNodes.Caption = "Applying the IPP exceptions - Purpose of disclosure"
        ElseIf 13 < nodeName < 34 Then
            fmNodes.Caption = "Applying the IPP exceptions - Maintenance of the law and related exceptions"
        ElseIf 33 < nodeName < 43 Then
            fmNodes.Caption = "Applying the IPP exceptions - Serious threat"
        ElseIf 50 < nodeName < 57 Then
            fmNodes.Caption = "Legislative vehicles for sharing personal information - Overview"
        End If
    End If

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
    Dim i As Integer
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
