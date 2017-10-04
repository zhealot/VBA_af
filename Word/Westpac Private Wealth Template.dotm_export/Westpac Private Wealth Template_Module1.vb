Attribute VB_Name = "Module1"
'****************************
'* Document bookmark rules: *
'****************************
'separated by '_';
'first part indicates this part exists in which plan: 'I' for individual, 'J' for joint and 'T' for Trust.
'second part indicate what section this book mark exists: 'Paragraph', 'Row' and 'Table'
'**************************************************
'string to hold Executive & Senior adviser names
Public Const ExceAdviser = "Johnathan Bayley,Karl Mabbutt,Quintin Budler,Richard McCadden,Simon Hepple,Sarah Priddle,Anna Francis -PWM Executive"
Public Const SeniorAdviser = "Johnathan Bayley,Karl Mabbutt,Quintin Budler,Richard McCadden,Simon Hepple,Sarah Priddle"
Public AccountType As String    'account type, 'I' for Individual, 'J' for Joint and 'T' for trust
Public KiwiSaver As Boolean     'whether Kiwisaver included
Public KSOnly As Boolean        'KiwiSaver only plan
Public PlanType As String       'plan type: 'A': Active, 'K': Kiwisaver only, 'W': Kiwisaver within Active
Public SchemeName  As String    'scheme name.
Public Scheme As String         'scheme type: 'DF': Default Fund; 'CF': Conservative Fund; 'MF': Moderate Fund; 'GF': Growth Fund; 'BF': Balanced Fund;
                                '             'B3': Belended 30/70; 'B5': Belended 50/50; 'B7':Belended 70/30.
Public sAdviser As String       'store adviser name
Public Const PassWord = "DavidPenny"

'************Bookmark notes***********************************************************************************
'
'*************************************************************************************************************


'****document varibles****
'ClientNames - 'known as' names of client, if not set, use formal names
'ClientNameFormal - formatl names of client
'TrustName - name of the truest
'Month - Month name, format mmmm
'AdviserName - Adviser name
'yourTrust's - assign to 'Trust's' for Trust, otherwise 'your', lower case
'youTrust - assign 'Trust' for Trust, otherwise 'you'
'yourTrust's_U - upper case
'youTrust_U- upper case
'youTrustees - you or Trustees
'youTrustees_U - upper case
'youtheir - you for individual, otherwise their
'youtheir_U -upper case
'yourits - your/its
'youthey - you/they
'PlanName - name of the chosen plan
'GG - percentage of Gain
'II - percentage of Income
'FundName1 - one of: Conservative/Moderate/Balanced/Growth Trust
'FundName2 - FundName2.'
'areis - are for 'you', is for 'Trust'
'ClientFormal1- client 1 formal name
'ClientFormal2 - client 2 formal name
'Client1 - client 1 known as name
'Client2 - client 2 known as name
'Beneficiaries - names of beneficiaries
'dodoes - do for individual, does for trust
'endings - "" for you, 's' for trust
'RiskProfile - Risk Profile
'
'

Function AccountClick(ob As MSForms.OptionButton)
'handles click action in account type frame, could be 'Individual/Joint/Trust
    Select Case ob.Name
    Case "obIndi"
    'Individual account
        AccountType = "I"
        CtrVisible frmMain.lbName1, True
        CtrVisible frmMain.lbName2, False
        CtrVisible frmMain.lbName3, False
        CtrVisible frmMain.lbName4, False
        CtrVisible frmMain.lbKnowAs1, True
        CtrVisible frmMain.lbKnowAs2, False
        CtrVisible frmMain.lbKnowAs3, False
        CtrVisible frmMain.lbKnowAs4, False
        CtrVisible frmMain.tbName1, True
        CtrVisible frmMain.tbName2, False
        CtrVisible frmMain.tbName3, False
        CtrVisible frmMain.tbName4, True
        CtrVisible frmMain.tbName5, False
        CtrVisible frmMain.tbName6, False
        CtrVisible frmMain.tbName7, False
        CtrVisible frmMain.tbName8, False
        frmMain.lbName1.Caption = "Client Name:"
        frmMain.lbKnowAs1.Caption = "Known As:"
        FrameAble frmMain.frmBene, False
        CtrVisible frmMain.obActive, True
        CtrVisible frmMain.obKiwi, True
        CtrVisible frmMain.obKiwiActive, True
    Case "obJoint"
    'Joint account
        AccountType = "J"
        CtrVisible frmMain.lbName1, True
        CtrVisible frmMain.lbName2, True
        CtrVisible frmMain.lbName3, False
        CtrVisible frmMain.lbName4, False
        CtrVisible frmMain.lbKnowAs1, True
        CtrVisible frmMain.lbKnowAs2, True
        CtrVisible frmMain.lbKnowAs3, False
        CtrVisible frmMain.lbKnowAs4, False
        CtrVisible frmMain.tbName1, True
        CtrVisible frmMain.tbName2, True
        CtrVisible frmMain.tbName3, False
        CtrVisible frmMain.tbName4, True
        CtrVisible frmMain.tbName5, True
        CtrVisible frmMain.tbName6, False
        CtrVisible frmMain.tbName7, False
        CtrVisible frmMain.tbName8, False
        frmMain.lbName1.Caption = "Client 1 Name:"
        frmMain.lbName2.Caption = "Client 2 Name:"
        frmMain.lbKnowAs1.Caption = "Known As:"
        frmMain.lbKnowAs2.Caption = "Known As:"
        FrameAble frmMain.frmBene, False
        CtrVisible frmMain.obActive, True
        CtrVisible frmMain.obKiwi, True
        CtrVisible frmMain.obKiwiActive, True
    Case "obTrust"
    'Trust account
        AccountType = "T"
        CtrVisible frmMain.lbName1, True
        CtrVisible frmMain.lbName2, True
        CtrVisible frmMain.lbName3, True
        CtrVisible frmMain.lbName4, True
        CtrVisible frmMain.lbKnowAs1, False
        CtrVisible frmMain.lbKnowAs2, True
        CtrVisible frmMain.lbKnowAs3, True
        CtrVisible frmMain.lbKnowAs4, True
        CtrVisible frmMain.tbName1, True
        CtrVisible frmMain.tbName2, True
        CtrVisible frmMain.tbName3, True
        CtrVisible frmMain.tbName4, False
        CtrVisible frmMain.tbName5, True
        CtrVisible frmMain.tbName6, True
        CtrVisible frmMain.tbName7, True
        CtrVisible frmMain.tbName8, True
        frmMain.lbName1.Caption = "Trust Name:"
        frmMain.lbName2.Caption = "Trustee Name:"
        frmMain.lbName3.Caption = "Trustee Name:"
        frmMain.lbName4.Caption = "Trustee Name:"
        frmMain.lbKnowAs1.Caption = ""
        frmMain.lbKnowAs2.Caption = "Known As:"
        frmMain.lbKnowAs3.Caption = "Known As:"
        frmMain.lbKnowAs4.Caption = "Known As:"
        FrameAble frmMain.frmBene, True
        CtrVisible frmMain.obActive, True
        CtrVisible frmMain.obKiwi, False
        CtrVisible frmMain.obKiwiActive, False
        frmMain.obActive.Value = True
    End Select
End Function

Function PlanClick(ob As MSForms.OptionButton)
'handles click on plan tyep option buttons.
    PlanType = ob.Name
    DocPrty "PlanName", "Westpac " & ob.Caption
    Select Case ob.Name
    Case "obActive"
        frmMain.frmScheme.Caption = "Westpac Active Series"
        frmMain.obDefault.Enabled = False
        KiwiSaver = False
        KSOnly = False
        DocPrty "KSOnly", "No"
    Case "obKiwi"
        frmMain.frmScheme.Caption = "Westpac KiwiSaver Scheme included with Active"
        frmMain.obDefault.Enabled = True
        KiwiSaver = True
        KSOnly = True
        DocPrty "KSOnly", "Yes"
    Case "obKiwiActive"
        frmMain.frmScheme.Caption = "Westpac KiwiSaver Scheme Only"
        frmMain.obDefault.Enabled = True
        KiwiSaver = True
        KSOnly = False
        DocPrty "KSOnly", "No"
    Case Else
    End Select
End Function

Function SchemeClick(ob As MSForms.OptionButton)
'handles click on scheme option buttons.
    Scheme = ob.Name
    SchemeName = ob.Caption
    'assign GG and II values
    'assign fund 1 name and fund 2 name
    Select Case Scheme
    Case "obDefault"
        DocPrty "GG", "20"
        DocPrty "II", "80"
        DocPrty "GG1", "20"
        DocPrty "II1", "80"
        DocPrty "FundName1", IIf(KSOnly, ob.Caption, Replace(ob.Caption, "Fund", "Trust"))  'for KS only plan, use 'Fund', otherwise 'Trust'
        DocPrty "FundName2", " "
        DocPrty "Blended", " "
    Case "obConservative"
        DocPrty "GG", "20"
        DocPrty "II", "80"
        DocPrty "GG1", "20"
        DocPrty "II1", "80"
        DocPrty "FundName1", IIf(KSOnly, ob.Caption, Replace(ob.Caption, "Fund", "Trust"))  'for KS only plan, use 'Fund', otherwise 'Trust'
        DocPrty "FundName2", " "
        DocPrty "Blended", " "
    Case "obModerate"
        DocPrty "GG", "40"
        DocPrty "II", "60"
        DocPrty "GG1", "40"
        DocPrty "II1", "60"
        DocPrty "FundName1", IIf(KSOnly, ob.Caption, Replace(ob.Caption, "Fund", "Trust"))  'for KS only plan, use 'Fund', otherwise 'Trust'
        DocPrty "FundName2", " "
        DocPrty "Blended", " "
    Case "obBalanced"
        DocPrty "GG", "60"
        DocPrty "II", "40"
        DocPrty "GG1", "60"
        DocPrty "II1", "40"
        DocPrty "FundName1", IIf(KSOnly, ob.Caption, Replace(ob.Caption, "Fund", "Trust"))  'for KS only plan, use 'Fund', otherwise 'Trust'
        DocPrty "FundName2", " "
        DocPrty "Blended", " "
    Case "obGrowth"
        DocPrty "GG", "80"
        DocPrty "II", "20"
        DocPrty "GG1", "80"
        DocPrty "II1", "20"
        DocPrty "FundName1", IIf(KSOnly, ob.Caption, Replace(ob.Caption, "Fund", "Trust"))  'for KS only plan, use 'Fund', otherwise 'Trust'
        DocPrty "FundName2", " "
        DocPrty "Blended", " "
    Case "obBlended37"
        DocPrty "GG", "30"
        DocPrty "II", "70"
        DocPrty "GG1", "20"
        DocPrty "II1", "80"
        DocPrty "GG2", "40"
        DocPrty "II2", "60"
        DocPrty "FundName1", "Conservative " & IIf(KSOnly, "Fund", "Trust")
        DocPrty "FundName2", "Moderate " & IIf(KSOnly, "Fund", "Trust")
        DocPrty "Blended", "Blended"
    Case "obBlended55"
        DocPrty "GG", "50"
        DocPrty "II", "50"
        DocPrty "GG1", "60"
        DocPrty "II1", "40"
        DocPrty "GG2", "40"
        DocPrty "II2", "60"
        DocPrty "FundName1", "Balanced " & IIf(KSOnly, "Fund", "Trust")
        DocPrty "FundName2", "Moderate " & IIf(KSOnly, "Fund", "Trust")
        DocPrty "Blended", "Blended"
    Case "obBlended73"
        DocPrty "GG", "70"
        DocPrty "II", "30"
        DocPrty "GG1", "60"
        DocPrty "II1", "40"
        DocPrty "GG2", "80"
        DocPrty "II2", "20"
        DocPrty "FundName1", "Balanced " & IIf(KSOnly, "Fund", "Trust")
        DocPrty "FundName2", "Growth " & IIf(KSOnly, "Fund", "Trust")
        DocPrty "Blended", "Blended"
    End Select
        DocPrty "KSFundName1", Replace(GetDocPrty("FundName1"), " Trust", " Fund") 'KiwiSaver part, always use 'Fund'
        DocPrty "KSFundName2", Replace(GetDocPrty("FundName2"), " Trust", " Fund") 'KiwiSaver part, always use 'Fund'
End Function

Function FrameAble(frm As MSForms.Frame, YesNo As Boolean)
'enable/disable frame and controls within
    frm.Enabled = YesNo
    For Each ctr In frm.Controls
        ctr.Enabled = YesNo
    Next ctr
End Function

Function CtrVisible(ctr As MSForms.control, YesNo As Boolean)
    ctr.Enabled = YesNo
    ctr.Visible = YesNo
End Function

Function DocPrty(nm As String, vl As String)
'set Custom Docuemnt Property
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(nm).Value = vl
End Function

Function GetDocPrty(nm As String) As String
'get Custom Document Property
    On Error Resume Next
    GetDocPrty = ActiveDocument.CustomDocumentProperties(nm).Value
End Function
Sub ReestVariables()
    For Each v In ActiveDocument.CustomDocumentProperties
        v.Value = v.Name
    Next v
    ActiveDocument.Content.Fields.Update
End Sub

Function DeleteViaBm(bm As Bookmark)
'delete document parts(paragraph/table/row/range) by bookmark
    'exit if no '_' char in bookmark name
    If InStr(bm.Name, "_") = 0 Then Exit Function
    
    Dim rg As Range
    Set rg = bm.Range
    'analysis bookmark name
    Dim AccType As String   'to store account type word, could be: I/J/T/K
    Dim Element As String  'to store element, could be: Paragraph/Talbe/Row/Range
    Dim arr() As String
    arr = Split(bm.Name, "_")
    AccType = arr(0)
    Element = arr(1)
    '(bookmark has a different account type, proceed delete)
    'Or (for those only shoudl exist in KiwiSaver and current Plan Type is Active)
    'Or (it's a KiwiSaver plan and bookmarked NonKiwiSaver)
    If bm.Name = "IJ_Paragraph_8" And PlanType = "obKiwi" Then '### an ugly exception
        Exit Function
    End If
    If InStr(AccType, AccountType) = 0 _
        Or (InStr(AccType, "K") > 0 And PlanType = "obActive") _
        Or (KiwiSaver And InStr(AccType, "NonKiwiSaver") > 0) _
        Or (InStr(AccType, "K") = 0 And PlanType = "obKiwi") Then
        'account type selected in pop up is different from account part in bookmark
        'Or: plan type is Active only but bookmakr has K(iwiSaver) in it
        'Or: KiwiSaver plan but bookmakred non-kiwisaver
        'Or: plantyep is K(iwiSaver) only but bookmark has no K init
        Select Case Element
        Case "Paragraph"
            If bm.Range.Paragraphs.Count > 0 Then
                bm.Range.Paragraphs(1).Range.Delete
            End If
            bm.Delete
        Case "Cell"
            If bm.Range.Tables.Count > 0 Then
                bm.Range.Cells(1).Delete
            End If
            bm.Delete
        Case "Row"
            If bm.Range.Tables.Count > 0 Then
                bm.Range.Rows.Delete
            End If
            bm.Delete
        Case "Table"
            If bm.Range.Tables.Count > 0 Then
                bm.Range.Tables(1).Delete
                bm.Range.Paragraphs(1).Range.Delete 'delete the paragraphs the table sits in, assuming all table in a stand alone paragraph
            End If
            bm.Delete
        Case "Text"
            bm.Range.Delete
        Case "Range"
            If InStr(bm.Name, "Start") > 0 Then
                If bm.Parent.Bookmarks.Exists(Replace(bm.Name, "Start", "End")) Then
                    rg.SetRange bm.Range.Start, bm.Parent.Bookmarks(Replace(bm.Name, "Start", "End")).Range.End
                    rg.Delete
                    bm.Range.Paragraphs(1).Range.Delete
                    bm.Parent.Bookmarks(Replace(bm.Name, "Start", "End")).Delete
                    bm.Delete
                End If
            End If
        Case "Column"
        'delete content in column, and merge each cell with previous column to make the tale same width
            If bm.Range.Tables.Count > 0 Then
                Dim clmIndex As Integer
                clmIndex = bm.Range.Cells(1).ColumnIndex    'column number
                If clmIndex < 2 Then Exit Function      'check table has more then 1 column
                Set rg = bm.Range
                Dim tb As Table
                Set tb = rg.Tables(1)
                On Error Resume Next
                For i = 1 To tb.Rows.Count
                    tb.Cell(i, clmIndex).Range.Text = ""
                    Set rg = tb.Cell(i, clmIndex - 1).Range                 'set range to previous cell
                    rg.SetRange rg.Start, tb.Cell(i, clmIndex).Range.End    'redifine range to contain two cells in the row
                    rg.Cells.Merge
                Next i
                bm.Delete
            End If
        Case "Bookmark"
            bm.Range.Text = ""
        Case Else
            Exit Function
        End Select
    End If
End Function

Sub testmerge()
    Dim rg As Range
    Dim rgTmp As Range
    Dim tb As Table
    Set rg = ThisDocument.Bookmarks("JT_Column_1").Range
    Set tb = rg.Tables(1)
    clmIndex = rg.Cells(1).ColumnIndex
    For i = 1 To tb.Rows.Count
        tb.Cell(i, clmIndex).Range.Text = ""
        Set rgTmp = tb.Cell(i, clmIndex - 1).Range
        rgTmp.SetRange rgTmp.Start, tb.Cell(i, clmIndex).Range.End
        rgTmp.Cells.Merge
    Next i
End Sub

Function LoadAdviser(cb As MSForms.ComboBox)
    Dim doc As Document
    Set doc = ThisDocument
    Dim bb As BuildingBlock
    Dim bbName As String
    Dim sName As String
    
    cb.Clear
    For i = 1 To doc.AttachedTemplate.BuildingBlockEntries.Count
        Set bb = doc.AttachedTemplate.BuildingBlockEntries(i)
        bbName = Trim(bb.Name)
        If Left(bb.Name, 3) = "AFA" And Right(bb.Name, 5) <> "SecDS" Then
            sName = Right(bb.Name, Len(bb.Name) - 3)
            'find out second capital letter to inset space between first and last name
            If Len(sName) > 1 Then
                For j = 1 To Len(sName) - 1
                    If Mid(sName, j + 1, 1) <= "Z" Then
                        sName = Left(sName, j) & " " & Right(sName, Len(sName) - j)
                        Exit For
                    End If
                Next j
                cb.AddItem sName
            End If
        End If
    Next i
End Function

'button visibility in ribbon
Sub ReturnVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IIf(ActiveDocument.Type = wdTypeTemplate, True, False)
End Sub



Sub Finalize(control As IRibbonControl)
'finalize document: key words replacement, ###lock down for protection
'[are/is]
'[do/does]
'[have/has]
'[you/it]
'[your/the Beneficiaries']
'[your/the Trust's]
'[Your/the Trust's]
'[you/the Trust]
'[You/the Trust]
'[you/The Trustees]
'[You/The Trustees]
'[your/their]
'[you/they]
'[You/They]
'[your/it's]
'###[wish/es]
'###[invest/s]
    If MsgBox("Check the document for replaceable words and lock down.", vbYesNo) = vbNo Then
        Exit Sub
    End If
    ReplaceWords "[are/is]", GetDocPrty("areis")
    ReplaceWords "[do/does]", GetDocPrty("dodoes")
    ReplaceWords "[have/has]", GetDocPrty("havehas")
    ReplaceWords "[you/it]", GetDocPrty("youit")
    ReplaceWords "[your/the Beneficiaries']", GetDocPrty("yourBeneficiaries")
    ReplaceWords "[your/the Trust's]", GetDocPrty("yourTrust's")
    ReplaceWords "[Your/the Trust's]", GetDocPrty("yourTrust's_U")
    ReplaceWords "[you/the Trust]", GetDocPrty("youTrust")
    ReplaceWords "[You/the Trust]", GetDocPrty("youTrust_U")
    ReplaceWords "[you/The Trustees]", GetDocPrty("youTrustees")
    ReplaceWords "[You/The Trustees]", GetDocPrty("youTrustees_U")
    ReplaceWords "[your/their]", GetDocPrty("youtheir")
    ReplaceWords "[you/they]", GetDocPrty("youthey")
    ReplaceWords "[You/They]", GetDocPrty("youthey_U")
    ReplaceWords "[your/it's]", GetDocPrty("yourits")
    'replace words like [show/s]
    'find pattern "\[*/s\]", user backslash "\" to escape special char
    ReplaceWithWildcard "/es", GetDocPrty("endinges")
    ReplaceWithWildcard "/s", GetDocPrty("endings")
    
    'update fields
    ActiveDocument.Fields.Update
End Sub

Function ReplaceWords(src As String, rpl As String)
    With ActiveDocument.Content.Find
        .ClearAllFuzzyOptions
        .Forward = True
        .MatchCase = True
        .Wrap = wdFindContinue
        .Text = src
        .Replacement.Text = rpl
        .Execute Replace:=wdReplaceAll
    End With
End Function

Function ReplaceWithWildcard(foo As String, bar As String)
'search text with "[", "]" and wildcard, replace part of it with new text
'foo:       "/es"
'bar:       ""
'e.g.:      [wish/es] -> wish

    'check sWhole has wildcard and foo in it.
    Dim ReplaceText As String
    Dim rg As Range
    Set rg = ActiveDocument.Content
    With rg.Find
        .ClearAllFuzzyOptions
        .Format = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Text = "\[*" & foo & "\]"
        .Execute
        Do While .Found
           ReplaceText = rg.Text
           ReplaceText = Replace(Replace(ReplaceText, "[", ""), "]", "")
           ReplaceText = Replace(ReplaceText, foo, bar)
           .Replacement.Text = ReplaceText
           .Execute Replace:=wdReplaceOne
           .Text = "\[*" & foo & "\]"
           rg.SetRange rg.End, ActiveDocument.Content.End
           .Execute
        Loop
    End With
End Function

'load Autotext from document
Sub UpdateAutotext(control As IRibbonControl)
    MsgBox "Please select Autotext document."
    Application.ScreenUpdating = False
    Dim sFilename As String
    'choose templates folder
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select Autotext document."
        .Filters.Clear
        .Filters.Add "Word document", "*.docx", 1
        .InitialFileName = ThisDocument.Path & "\"
        .InitialView = msoFileDialogViewDetails
        .ButtonName = "OK"
        If .Show = -1 Then
            sFilename = .SelectedItems(1)
        End If
    End With
    'user pressed 'Cancel'
    If sFilename = "" Then End
    TransferBB sFilename, False
    Application.ScreenUpdating = True
End Sub

Function TransferBB(sSrc As String, fromBB As Boolean)
'function to add building blocks into this document
'fromBB = true, add from source document's attached template's Building Blocks
'fromBB = false, add from source document content
    Dim srcDoc As Document
    Dim sPrefix As String
    Dim sEnding As String
    Dim bbName As String
    Dim bbCategory As String
    Dim Msg As String
    Dim rgStart As Long
    Dim rgEnd As Long
    Dim rgTmp As Range
    Dim iCounter As Integer
    Dim rg As Range
    
    sPrefix = "%%"
    sEnding = "<<EndAutotext>>"
    
    Application.ScreenUpdating = False   '###
    'delete existing building blocks
    While ThisDocument.AttachedTemplate.BuildingBlockEntries.Count > 0
        ThisDocument.AttachedTemplate.BuildingBlockEntries(1).Delete
    Wend
    
    On Error Resume Next
    Set srcDoc = Documents.Open(sSrc, ReadOnly:=True, Visible:=False)
    If Err.Number <> 0 Then
        Exit Function
    End If
    If fromBB Then
    'add BB from source document's building blocks
        Dim tmpDoc As Document
        Set tmpDoc = Documents.Add(, , , False)  '###
        Set dstDoc = ThisDocument
        For i = 1 To srcDoc.AttachedTemplate.BuildingBlockEntries.Count
            tmpDoc.Content.Delete
            Set rg = tmpDoc.Content
            rg.Collapse wdCollapseStart
            srcDoc.AttachedTemplate.BuildingBlockEntries(i).Insert rg, True
            Set rg = tmpDoc.Content
            bbName = srcDoc.AttachedTemplate.BuildingBlockEntries(i).Name
            bbCategory = "General"
            ThisDocument.AttachedTemplate.BuildingBlockEntries.Add Name:=bbName, _
                                    Type:=wdTypeAutoText, Category:=bbCategory, Range:=rg, InsertOptions:=wdInsertContent
        Next i
        tmpDoc.Close False
        Set tmpDoc = Nothing
        Msg = i & " Building block items have been added."
    Else
    'add from document content
        Set rgTmp = srcDoc.Content
        With rgTmp
            .Find.ClearAllFuzzyOptions
            .Find.Text = sPrefix
            .Find.Wrap = wdFindStop
            .Find.Forward = True
            .Find.Execute
            Do While .Find.Found
                If InStr(2, rgTmp.Paragraphs(1).Range.Text, sPrefix) > 0 Then
                'check if the paragraphs of the '%%XXXX%%' format
                    bbName = Replace(rgTmp.Paragraphs(1).Range.Text, sPrefix, "")
                    rgStart = rgTmp.Next(wdParagraph, 1).Start
                    rgTmp.SetRange rgStart, srcDoc.Content.End
                    rgTmp.Find.Text = sEnding
                    rgTmp.Find.Execute
                    If rgTmp.Find.Found Then
                        rgEnd = rgTmp.Previous(wdParagraph, 1).End 'previous paragraph's end
                        rgTmp.SetRange rgStart, rgEnd
                        ThisDocument.AttachedTemplate.BuildingBlockEntries.Add Name:=bbName, _
                                            Type:=wdTypeAutoText, Category:="General", Range:=rgTmp, InsertOptions:=wdInsertContent
                        iCounter = iCounter + 1
                    End If
                End If
                If rgTmp.End < srcDoc.Content.End Then
                    rgTmp.SetRange rgTmp.End, srcDoc.Content.End
                End If
                .Find.Text = sPrefix
                .Find.Execute
            Loop
        End With
        Msg = iCounter & " building blocks added."
    End If
    srcDoc.Close False
    Set srcDoc = Nothing
    Application.ScreenUpdating = True
    MsgBox Msg
End Function

