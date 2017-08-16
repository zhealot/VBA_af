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
    If frmMain.cbExeAdviser.Value = "" Then
        MsgBox "Please choose an Executive Adviser."
        frmMain.cbExeAdviser.SetFocus
        Exit Sub
    End If
    If frmMain.cbSeniorAdviser.Value = "" Then
        MsgBox "Please choose a Senior Adviser."
        frmMain.cbSeniorAdviser.SetFocus
        Exit Sub
    End If
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

    DocPrty "Month", Format(Date, "mmmm")
    DocPrty "AdviserName", cbAdviser.Value
    Select Case AccountType
    Case "I"    'individual account
        DocPrty "ClientNames", frmMain.tbName1.Text
        DocPrty "yourTrust's", "your"
        DocPrty "youTrust", "you"
        DocPrty "yourTrust's_U", "Your"
        DocPrty "youTrust_U", "You"
        DocPrty "youTrustees", "you"
        DocPrty "youTrustees_U", "You"
        DocPrty "youtheir", "you"
        DocPrty "youtheir_U", "You"
        DocPrty "yourits", "your"
        DocPrty "youthey", "you"
        DocPrty "areis", "are"
        DocPrty "dodoes", "do"
        DocPrty "endings", ""
        DocPrty "ClientFormal1", tbName1.Text
        DocPrty "ClientFormal2", " "
        DocPrty "Client1", IIf(tbName4.Text = "", tbName1.Text, tbName4.Text)
        DocPrty "Client2", " "
    Case "J"    'Joint account
        DocPrty "ClientNames", frmMain.tbName1.Text & " and " & frmMain.tbName2.Text
        DocPrty "yourTrust's", "your"
        DocPrty "youTrust", "you"
        DocPrty "yourTrust's_U", "Your"
        DocPrty "youTrust_U", "You"
        DocPrty "youTrustees", "you"
        DocPrty "youTrustees_U", "You"
        DocPrty "youtheir", "you"
        DocPrty "youtheir_U", "You"
        DocPrty "yourits", "your"
        DocPrty "youthey", "you"
        DocPrty "areis", "are"
        DocPrty "dodoes", "do"
        DocPrty "endings", ""
        DocPrty "ClientFormal1", tbName1.Text
        DocPrty "ClientFormal2", tbName2.Text
        DocPrty "Client1", IIf(tbName4.Text = "", tbName1.Text, tbName4.Text)
        DocPrty "Client2", IIf(tbName5.Text = "", tbName2.Text, tbName5.Text)
    Case "T"    'Trust account
        DocPrty "ClientNames", frmMain.tbName2.Text & " and " & frmMain.tbName3.Text
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
        DocPrty "endings", "s"
        DocPrty "TrustName", tbName1.Value
        DocPrty "ClientFormal1", tbName2.Text
        DocPrty "ClientFormal2", tbName3.Text
        DocPrty "Client1", IIf(tbName5.Text = "", tbName2.Text, tbName5.Text)
        DocPrty "Client2", IIf(tbName6.Text = "", tbName3.Text, tbName6.Text)
        DocPrty "Beneficiaries", tbBene1.Value & IIf(tbBene2.Value <> "", ", " & tbBene2.Value, "") & IIf(tbBene3.Value <> "", ", " & tbBene3.Value, "") & IIf(tbBene4.Value <> "", ", " & tbBene4.Value, "")
    End Select
    Me.Hide
    For Each v In ThisDocument.CustomDocumentProperties
        Debug.Print v.Name & " : " & v.Value
    Next v
    'delete non-relevant parts by bookmarks
    Dim bm As Bookmark
    Dim doc As Document
    Set doc = ActiveDocument
    For Each bm In doc
        
    Next bm
    ActiveDocument.Fields.Update
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
    Me.lbKnowAs1.Caption = ""
    Me.lbKnowAs2.Caption = ""
    Me.lbKnowAs3.Caption = ""
    Me.lbName1.Enabled = False
    Me.lbName2.Enabled = False
    Me.lbName3.Enabled = False
    Me.tbName1.Enabled = False
    Me.tbName2.Enabled = False
    Me.tbName3.Enabled = False
End Sub

