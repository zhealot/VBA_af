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
Public PlanType As String       'plan type: 'A': Active, 'K': Kiwisaver only, 'W': Kiwisaver within Active
Public Scheme As String         'scheme type: 'DF': Default Fund; 'CF': Conservative Fund; 'MF': Moderate Fund; 'GF': Growth Fund; 'BF': Balanced Fund;
                                '             'B3': Belended 30/70; 'B5': Belended 50/50; 'B7':Belended 70/30.
'************Bookmark notes***********************************************************************************
'Insert_1: to insert 'KIWISAVER' or 'WESTPAC ACTIVE'
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
        CtrVisible frmMain.lbKnowAs1, True
        CtrVisible frmMain.lbKnowAs2, False
        CtrVisible frmMain.lbKnowAs3, False
        CtrVisible frmMain.tbName1, True
        CtrVisible frmMain.tbName2, False
        CtrVisible frmMain.tbName3, False
        CtrVisible frmMain.tbName4, True
        CtrVisible frmMain.tbName5, False
        CtrVisible frmMain.tbName6, False
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
        CtrVisible frmMain.lbKnowAs1, True
        CtrVisible frmMain.lbKnowAs2, True
        CtrVisible frmMain.lbKnowAs3, False
        CtrVisible frmMain.tbName1, True
        CtrVisible frmMain.tbName2, True
        CtrVisible frmMain.tbName3, False
        CtrVisible frmMain.tbName4, True
        CtrVisible frmMain.tbName5, True
        CtrVisible frmMain.tbName6, False
        frmMain.lbName1.Caption = "Client 1 Name:"
        frmMain.lbName2.Caption = "Cleint 2 Name:"
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
        CtrVisible frmMain.lbKnowAs1, False
        CtrVisible frmMain.lbKnowAs2, True
        CtrVisible frmMain.lbKnowAs3, True
        CtrVisible frmMain.tbName1, True
        CtrVisible frmMain.tbName2, True
        CtrVisible frmMain.tbName3, True
        CtrVisible frmMain.tbName4, False
        CtrVisible frmMain.tbName5, True
        CtrVisible frmMain.tbName6, True
        frmMain.lbName1.Caption = "Trust Name:"
        frmMain.lbName2.Caption = "Trustee Name:"
        frmMain.lbName3.Caption = "Trustee Name:"
        frmMain.lbKnowAs1.Caption = ""
        frmMain.lbKnowAs2.Caption = "Known As:"
        frmMain.lbKnowAs3.Caption = "Known As:"
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
    DocPrty "PlanName", ob.Caption
    Select Case ob.Name
    Case "obActive"
        frmMain.frmScheme.Caption = "Westpac Active Series"
        frmMain.obDefault.Enabled = False
    Case "obKiwi"
        frmMain.frmScheme.Caption = "Westpac KiwiSaver Scheme included with Active"
        frmMain.obDefault.Enabled = True
    Case "obKiwiActive"
        frmMain.frmScheme.Caption = "Westpac KiwiSaver Scheme Only"
        frmMain.obDefault.Enabled = True
    Case Else
    End Select
End Function

Function SchemeClick(ob As MSForms.OptionButton)
'handles click on scheme option buttons.
    Scheme = ob.Name
    'assign GG and II values
    'assign fund 1 name and fund 2 name
    Select Case Scheme
    Case "obDefault"
        DocPrty "GG", "20"
        DocPrty "II", "80"
        DocPrty "FundName1", ob.Caption
        DocPrty "FundName2", " "
    Case "obConservative"
        DocPrty "GG", "20"
        DocPrty "II", "80"
        DocPrty "FundName1", ob.Caption
        DocPrty "FundName2", " "
    Case "obModerate"
        DocPrty "GG", "40"
        DocPrty "II", "60"
        DocPrty "FundName1", ob.Caption
        DocPrty "FundName2", " "
    Case "obBalanced"
        DocPrty "GG", "60"
        DocPrty "II", "40"
        DocPrty "FundName1", ob.Caption
        DocPrty "FundName2", " "
    Case "obGrowth"
        DocPrty "GG", "80"
        DocPrty "II", "20"
        DocPrty "FundName1", ob.Caption
        DocPrty "FundName2", " "
    Case "obBlended37"
        DocPrty "GG", "30"
        DocPrty "II", "70"
        DocPrty "FundName1", "Conservative Fund"
        DocPrty "FundName2", "Moderate Fund"
    Case "obBlended55"
        DocPrty "GG", "50"
        DocPrty "II", "50"
        DocPrty "FundName1", "Balanced Fund"
        DocPrty "FundName2", "Moderate Fund"
    Case "obBlended73"
        DocPrty "GG", "70"
        DocPrty "II", "30"
        DocPrty "FundName1", "Balanced Fund"
        DocPrty "FundName2", "Growth Fund"
    End Select
End Function

Function FrameAble(frm As MSForms.Frame, YesNo As Boolean)
'enable/disable frame and controls within
    frm.Enabled = YesNo
    For Each ctr In frm.Controls
        ctr.Enabled = YesNo
    Next ctr
End Function

Function CtrVisible(ctr As MSForms.Control, YesNo As Boolean)
    ctr.Enabled = YesNo
    ctr.Visible = YesNo
End Function

Function DocPrty(nm As String, vl As String)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(nm).Value = vl
End Function

Sub docpry()
'    For Each v In ThisDocument.CustomDocumentProperties
'        Debug.Print v.Name
'    Next v
    Dim arr() As String
    arr = Split("IJ_Row_3", "_")
    Debug.Print arr(0)
    Debug.Print arr(1)
    
End Sub

Sub ReestVariables()
    For Each v In ActiveDocument.CustomDocumentProperties
        v.Value = v.Name
    Next v
    ActiveDocument.Content.Fields.Update
End Sub

Function DeleteViaBm(bm As Bookmark)
'delete document parts(paragraph/table/row/range) by bookmark
    'If Not doc.Bookmarks.Exists(bm.Name) Then Exit Function
    Dim rg As Range
    Dim doc As Document
    Set doc = bm.Parent
    Set rg = bm.Range
    'analysis bookmark name
    Dim AccType As String   'to store account type word, could be: I/J/T/K
    Dim Element As String  'to store element, could be: Paragraph/Talbe/Row/Range
    Dim arr() As String
    arr = Split(bm.Name, "_")
    AccType = arr(0)
    Element = arr(1)
    'bookmark has a different account type, proceed delete
    If InStr(AccType, AccountType) = 0 Then
        Select Case Element
        Case "Paragraph"
            Set rg = bm.Range.Paragraphs(1).Range
        Case "Row"
            Set rg = bm.Range.Rows(1).Range
        Case "Table"
            Set rg = bm.Range.Tables(1).Range
        End Select
        rg.Delete
    End If
    'for those only shoudl exist in KiwiSaver and current Plan Type is Active
    If InStr(AccType, "K") And PlanType = "A" Then
        
    End If
End Function


