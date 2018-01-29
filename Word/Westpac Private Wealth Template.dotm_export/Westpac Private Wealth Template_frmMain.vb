VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Westpac Private Wealth Financial Plan"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   OleObjectBlob   =   "Westpac Private Wealth Template_frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Dim aryTmp() As String
    'populate Executive Adviser dropdown
    aryTmp = Split(ExceAdviser, ",")
    For Each v In aryTmp
        cbExeAdviser.AddItem v
    Next v
    'populate Senior Adviser dropdown
    aryTmp = Split(SeniorAdviser, ",")
    For Each v In aryTmp
        cbSeniorAdviser.AddItem v
    Next v
    Me.lbName1.Caption = ""
    Me.lbName2.Caption = ""
    Me.lbName3.Caption = ""
    Me.lbName4.Caption = ""
    Me.lbKnowAs1.Caption = ""
    Me.lbKnowAs2.Caption = ""
    Me.lbKnowAs3.Caption = ""
    Me.lbKnowAs4.Caption = ""
    Me.lbName1.Enabled = False
    Me.lbName2.Enabled = False
    Me.lbName3.Enabled = False
    Me.lbName4.Enabled = False
    Me.tbName1.Enabled = False
    Me.tbName2.Enabled = False
    Me.tbName3.Enabled = False
    Me.tbName4.Enabled = False
    LoadAdviser cbAdviser
End Sub

Private Sub btnOK_Click()
    'check necessaries
    If AccountType = "" Then
        MsgBox "Please choose an account type."
        Exit Sub
    End If
    If PlanType = "" Then
        MsgBox "Please choose a plan type."
        Exit Sub
    End If
    If Scheme = "" Then
        MsgBox "Please choose a scheme."
        Exit Sub
    End If
    'check by account type
    Select Case AccountType
    Case "I"
    'Individual
        If Trim(frmMain.tbName1) = "" Then
            MsgBox "Please enter client name."
            frmMain.tbName1.SetFocus
            Exit Sub
        End If
    Case "J"
    'Joint
        If Trim(frmMain.tbName1) = "" Then
            MsgBox "Please enter client 1 name."
            frmMain.tbName1.SetFocus
            Exit Sub
        End If
        If Trim(frmMain.tbName2) = "" Then
            MsgBox "Please enter client 2 names."
            frmMain.tbName2.SetFocus
            Exit Sub
        End If
    Case "T"
    'Trust
        If Trim(frmMain.tbName1) = "" Then
            MsgBox "Please enter Trust name."
            frmMain.tbName1.SetFocus
            Exit Sub
        End If
        If Trim(frmMain.tbName2) = "" Then
            MsgBox "Please enter Trustee 1 name."
            frmMain.tbName2.SetFocus
            Exit Sub
        End If
        If Trim(frmMain.tbName3) = "" Then
            MsgBox "Please enter Trustee 2 name."
            frmMain.tbName3.SetFocus
            Exit Sub
        End If
        If Trim(frmMain.tbBene1) = "" Then
            MsgBox "Please enter beneficiary names."
            frmMain.tbBene1.SetFocus
            Exit Sub
        End If
    End Select
    'check Advisers
    If frmMain.cbAdviser.Value = "" Then
        MsgBox "Please choose an Adviser."
        frmMain.cbAdviser.SetFocus
        Exit Sub
    End If
'    If frmMain.cbExeAdviser.Value = "" Then
'        MsgBox "Please choose an Executive Adviser."
'        frmMain.cbExeAdviser.SetFocus
'        Exit Sub
'    End If
'    If frmMain.cbSeniorAdviser.Value = "" Then
'        MsgBox "Please choose a Senior Adviser."
'        frmMain.cbSeniorAdviser.SetFocus
'        Exit Sub
'    End If
    'check docu no.
    If Trim(frmMain.tbCRS) = "" Then
        MsgBox "Please enter a CRS number."
        frmMain.tbCRS.SetFocus
        Exit Sub
    End If
    If Trim(frmMain.tbFPW) = "" Then
        MsgBox "Please enter a FPW number."
        frmMain.tbFPW.SetFocus
        Exit Sub
    End If
    
    'get Adviser title
    DocPrty "AdviserTitle", "Senior Financial Adviser"
    Dim sAry() As String
    sAry = Split(SeniorAdviser, ",")
    For Each vr In sAry
        If InStr(vr, cbAdviser.Value) > 0 Then
            DocPrty "AdviserTitle", "Executive Financial Adviser"
            Exit For
        End If
    Next vr
    'except for Anna Francis
    If InStr(cbAdviser.Value, "Anna Francis") > 0 Then
        DocPrty "AdviserTitle", "PWM Executive"
    End If
    Dim rg As Range
    Dim rgTmp As Range
    Dim doc As Document
    Dim bm As Bookmark
    Set doc = ActiveDocument
    Me.Hide
    
    'unlock template first
    If doc.ProtectionType <> wdNoProtection Then
        doc.Unprotect PassWord
        'incorrect password
        If doc.ProtectionType <> wdNoProtection Then
            MsgBox "Failed to unlock template, unable to proceed."
        End If
    End If
    
    DocPrty "date", Format(Date, "dd mmmm yyyy")
    DocPrty "AdviserName", cbAdviser.Value
    DocPrty "AdviserFirstName", Left(cbAdviser.Value, InStr(cbAdviser.Value, " ") - 1)
    DocPrty "FPW_Number", tbFPW.Value
    DocPrty "CRS_Number", tbCRS.Value
    Select Case AccountType
    Case "I"    'individual account
        DocPrty "ClientNames", IIf(tbName4.Text = "", tbName1.Text, tbName4.Text)
        DocPrty "ClientNameFormal", tbName1.Text
        DocPrty "ClientFormal1", tbName1.Text
        DocPrty "ClientFormal2", " "
        DocPrty "ClientFormal3", " "
        DocPrty "Client1", IIf(tbName4.Text = "", tbName1.Text, tbName4.Text)
        DocPrty "Client2", " "
        DocPrty "Client3", " "
        DocPrty "yourTrust's", "your"
        DocPrty "youTrust", "you"
        DocPrty "yourTrust's_U", "Your"
        DocPrty "youTrust_U", "You"
        DocPrty "youTrustees", "you"
        DocPrty "youTrustees_U", "You"
        DocPrty "youtheir", "your"
        DocPrty "youtheir_U", "Your"
        DocPrty "yourits", "your"
        DocPrty "youthey", "you"
        DocPrty "areis", "are"
        DocPrty "dodoes", "do"
        DocPrty "endings", " "
        DocPrty "endinges", " "
        DocPrty "TrustName", " "
        DocPrty "yourBeneficiaries", "your"
        DocPrty "havehas", "have"
        DocPrty "youit", "you"
        DocPrty "PersonalFinancial", "personal/financial"
        DocPrty "personal", " personal "
        DocPrty "IWe", "I"
        DocPrty "IWe_U", "I"
        DocPrty "AccountType", "I"
        DocPrty "myour", "my"
        DocPrty "amare", "am"
        DocPrty "yourTrustees", "your"
        DocPrty "yourTrustees_U", "Your"
        DocPrty "amare", "am"
        DocPrty "meus", "me"
        DocPrty "yoursTrustees", "yours"
        DocPrty "youBeneficiaries", "you"
        DocPrty "youthem", "you"
    Case "J"    'Joint account
        DocPrty "ClientNames", IIf(tbName4.Text = "", tbName1.Text, tbName4.Text) & " and " & IIf(tbName5.Text = "", tbName2.Text, tbName5.Text)
        DocPrty "ClientNameFormal", tbName1.Text & " and " & tbName2.Text
        DocPrty "ClientFormal1", tbName1.Text
        DocPrty "ClientFormal2", tbName2.Text
        DocPrty "ClientFormal3", " "
        DocPrty "Client1", IIf(tbName4.Text = "", tbName1.Text, tbName4.Text)
        DocPrty "Client2", IIf(tbName5.Text = "", tbName2.Text, tbName5.Text)
        DocPrty "Client3", " "
        DocPrty "yourTrust's", "your"
        DocPrty "youTrust", "you"
        DocPrty "yourTrust's_U", "Your"
        DocPrty "youTrust_U", "You"
        DocPrty "youTrustees", "you"
        DocPrty "youTrustees_U", "You"
        DocPrty "youtheir", "your"
        DocPrty "youtheir_U", "You"
        DocPrty "yourits", "your"
        DocPrty "youthey", "you"
        DocPrty "areis", "are"
        DocPrty "dodoes", "do"
        DocPrty "endings", " "
        DocPrty "endinges", " "
        DocPrty "TrustName", " "
        DocPrty "yourBeneficiaries", "your"
        DocPrty "havehas", "have"
        DocPrty "youit", "you"
        DocPrty "PersonalFinancial", "personal/financial"
        DocPrty "personal", " personal "
        DocPrty "IWe", "we"
        DocPrty "IWe_U", "We"
        DocPrty "AccountType", "J"
        DocPrty "myour", "our"
        DocPrty "amare", "are"
        DocPrty "yourTrustees", "your"
        DocPrty "yourTrustees_U", "Your"
        DocPrty "amare", "are"
        DocPrty "meus", "us"
        DocPrty "yoursTrustees", "yours"
        DocPrty "youBeneficiaries", "you"
        DocPrty "youthem", "you"
    Case "T"    'Trust account
        DocPrty "ClientNames", IIf(tbName5.Text = "", tbName2.Text, tbName5.Text) & IIf(tbName7.Text = "", " and ", ", ") & IIf(tbName6.Text = "", tbName3.Text, tbName6.Text) & IIf(tbName7.Text = "", "", " and " & IIf(tbName8.Text = "", tbName7.Text, tbName8.Text))
        DocPrty "ClientNameFormal", tbName2.Text & IIf(tbName7.Text = "", " and ", ", ") & tbName3.Text & IIf(tbName7.Text = "", "", " and " & tbName7)
        DocPrty "ClientFormal1", tbName2.Text
        DocPrty "ClientFormal2", tbName3.Text
        DocPrty "ClientFormal3", IIf(tbName7.Text = "", " ", tbName7.Text)
        DocPrty "Client1", IIf(tbName5.Text = "", tbName2.Text, tbName5.Text)
        DocPrty "Client2", IIf(tbName6.Text = "", tbName3.Text, tbName6.Text)
        DocPrty "Client3", IIf(tbName8.Text = "", tbName7.Text, tbName8.Text)
        DocPrty "yourTrust's", "the Trust's"
        DocPrty "youTrust", "the Trust"
        DocPrty "yourTrust's_U", "The Trust's"
        DocPrty "youTrust_U", "The Trust"
        DocPrty "youTrustees", "the Trustees"
        DocPrty "youTrustees_U", "The Trustees"
        DocPrty "youtheir", "their"
        DocPrty "youtheir_U", "Their"
        DocPrty "yourits", "its"
        DocPrty "youthey", "they"
        DocPrty "areis", "is"
        DocPrty "dodoes", "does"
        DocPrty "endings", "s "
        DocPrty "endinges", "es "
        DocPrty "TrustName", tbName1.Value
        DocPrty "Beneficiaries", tbBene1.Value & IIf(tbBene2.Value <> "", ", " & tbBene2.Value, "") & IIf(tbBene3.Value <> "", ", " & tbBene3.Value, "") & IIf(tbBene4.Value <> "", ", " & tbBene4.Value, "")
        DocPrty "yourBeneficiaries", "the Beneficiaries'"
        DocPrty "havehas", "has"
        DocPrty "youit", "it"
        DocPrty "PersonalFinancial", "financial"
        DocPrty "personal", " "
        DocPrty "IWe", "we"
        DocPrty "IWe_U", "We"
        DocPrty "AccountType", "T"
        DocPrty "myour", "our"
        DocPrty "amare", "are"
        DocPrty "yourTrustees", "the Trustees'"
        DocPrty "yourTrustees_U", "The Trustees'"
        DocPrty "amare", "are"
        DocPrty "meus", "us"
        DocPrty "yoursTrustees", "the Trustees'"
        DocPrty "youBeneficiaries", "the Beneficiaries"
        DocPrty "youthem", "them"
    End Select
    
    'write client names to bookmarked places
    If doc.Bookmarks.Exists("ClientName1") Then
        doc.Bookmarks("ClientName1").Range.Text = GetDocPrty("ClientNames")
    End If
    'write client formal names
    DocPrty "ClientNameFormal", GetDocPrty("ClientFormal1") & IIf(GetDocPrty("ClientFormal2") = " ", "", IIf(GetDocPrty("ClientFormal3") = " ", " and " & GetDocPrty("ClientFormal2"), ", " & GetDocPrty("ClientFormal2") & " and " & GetDocPrty("ClientFormal3")))
    'insert Adviser profile and reg
    If doc.Bookmarks.Exists("AdvisorDetail") Then
        Set rg = doc.Bookmarks("AdvisorDetail").Range
        rg.Collapse wdCollapseStart
        On Error Resume Next
        doc.AttachedTemplate.BuildingBlockEntries("AFA" & Replace(sAdviser, " ", "", , 1)).Insert rg, True  'replace onece 'cause there might be trailing space in BB name.
    End If
    If doc.Bookmarks.Exists("AdvisorReg") Then
        Set rg = doc.Bookmarks("AdvisorReg").Range
        rg.Collapse wdCollapseStart
        On Error Resume Next
        Set rgTmp = doc.AttachedTemplate.BuildingBlockEntries("AFA" & Replace(sAdviser, " ", "") & "SecDS").Insert(rg, True)
        'move adviser's name and reg text to above paragraph, delete them from here
        If rgTmp.Paragraphs.Count > 0 Then
            If rgTmp.Paragraphs(1).Range.Tables.Count = 0 Then
            'first para not in table
                rgTmp.Paragraphs(1).Range.Delete
                If rgTmp.Tables.Count > 0 Then
                    If doc.Bookmarks.Exists("AdvierNameReg") Then
                    'move text in first cell to here
                        doc.Bookmarks("AdvierNameReg").Range.Text = Replace(Trim(Left(rgTmp.Tables(1).Cell(1, 1).Range.Text, Len(rgTmp.Tables(1).Cell(1, 1).Range.Text) - 2)), Chr(13), "")
                        rgTmp.Tables(1).Rows(1).Delete
                        rgTmp.Paragraphs(1).Range.Delete
                    End If
                End If
            End If
        End If
    End If
    
    'delete other Adviser profile from Building Blocks
    Dim i As Integer
    For i = doc.AttachedTemplate.BuildingBlockEntries.Count To 1 Step -1
        If Left(doc.AttachedTemplate.BuildingBlockEntries(i).Name, 3) = "AFA" Then
            doc.AttachedTemplate.BuildingBlockEntries(i).Delete
        End If
    Next i
    '### after deleting building blocks, DO NOT save AttachedTempalte(this) after document is generated.
    ThisDocument.Saved = True
    
    'insert beneficiraries
    For Each bm In doc.Bookmarks
        If InStr(bm.Name, "Beneficiaries") > 0 Then
            bm.Range.Text = doc.CustomDocumentProperties("Beneficiaries")
        End If
    Next bm
    
    Dim sTmp As String
    'insert PortAct Autotext
    If doc.Bookmarks.Exists("PortAct1") Then
        If InStr(Scheme, "Blended") > 0 Then
            sTmp = Replace("Port" & IIf(KSOnly, "KS", "Act") & Left(SchemeName, InStr(SchemeName, "-") - 2), " ", "")
        Else
            sTmp = "Port" & IIf(KSOnly, "KS", "Act") & Left(SchemeName, InStr(SchemeName, " ") - 1)
        End If
        On Error Resume Next
        Set rg = doc.Bookmarks("PortAct1").Range
        Set rgTmp = doc.AttachedTemplate.BuildingBlockEntries(sTmp).Insert(rg, True)
    End If
    
    'insert Act Port autotext
    If doc.Bookmarks.Exists("ActPortOut") Then
        sTmp = "Act" & Left(SchemeName, 3) & "PortOut"
        On Error Resume Next
        Set rg = doc.Bookmarks("ActPortOut").Range
        Set rgTmp = doc.AttachedTemplate.BuildingBlockEntries(sTmp).Insert(rg, True)
    End If
    
    'delete rows in 'Financial Projections' table for non-blended scheme
    If InStr(SchemeName, "Blended") = 0 Then
        For i = 1 To 4
            If doc.Bookmarks.Exists("BlendedRow" & i) Then
                Set rg = doc.Bookmarks("BlendedRow" & i).Range
                doc.Bookmarks("BlendedRow" & i).Delete
                rg.Rows(1).Delete
            End If
        Next i
    End If
    
    'delete non-relevant parts by bookmarks
    For Each bm In doc.Bookmarks
        DeleteViaBm bm
    Next bm
    'if Trust plan and no 3rd trustee, then delete relevant rows in signature part
    If AccountType <> "T" Or tbName7.Text = "" Then
        Set bm = doc.Bookmarks("JT_Row_3")
        If bm.Range.Tables.Count > 0 Then
            bm.Range.Rows.Delete
        End If
        bm.Delete
        Set bm = doc.Bookmarks("JT_Row_4")
        If bm.Range.Tables.Count > 0 Then
            bm.Range.Rows.Delete
        End If
        bm.Delete
    End If
    ActiveDocument.Fields.Update
    ActiveDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Fields.Update
    If doc.ProtectionType = wdNoProtection Then
        doc.Protect wdAllowOnlyReading, , PassWord
    End If
End Sub

Private Sub cbAdviser_Change()
    sAdviser = cbAdviser.Value
End Sub

Private Sub obActive_Click()
    Call PlanClick(obActive)
End Sub

Private Sub obBalanced_Click()
    Call SchemeClick(obBalanced)
End Sub

Private Sub obBlended37_Click()
    Call SchemeClick(obBlended37)
End Sub

Private Sub obBlended55_Click()
    Call SchemeClick(obBlended55)
End Sub

Private Sub obBlended73_Click()
    Call SchemeClick(obBlended73)
End Sub

Private Sub obConservative_Click()
    Call SchemeClick(obConservative)
End Sub

Private Sub obDefault_Click()
    Call SchemeClick(obDefault)
End Sub

Private Sub obGrowth_Click()
    Call SchemeClick(obGrowth)
End Sub

Private Sub obIndi_Click()
    Call AccountClick(obIndi)
End Sub

Private Sub obJoint_Click()
    Call AccountClick(obJoint)
End Sub

Private Sub obKiwi_Click()
    Call PlanClick(obKiwi)
End Sub

Private Sub obKiwiActive_Click()
    Call PlanClick(obKiwiActive)
End Sub

Private Sub obModerate_Click()
    Call SchemeClick(obModerate)
End Sub

Private Sub obTrust_Click()
    Call AccountClick(obTrust)
End Sub


