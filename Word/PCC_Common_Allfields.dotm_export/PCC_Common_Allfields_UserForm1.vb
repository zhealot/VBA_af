VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "User Setup "
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   OleObjectBlob   =   "PCC_Common_Allfields_UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This common template was developed by Allfields
' Version 4.1, 5 July 2012, Shiree Hart - Allfields www.allfields.co.nz - Ph 04 978 7101 or email shiree@allfields.co.nz

' It is the supporting template with shared code
' It also seeks information from users to store in an
' inifile called user.ini that is stored in the folder determine by the public constant in the template module
' New users of Word should find they are prompted to fill in this
' information on first accessing Word
' This template should be stored in the nominated Startup folder that is typically (for convenience)
' stored in a subfolder of the Workgroup Templates folder.
' For templates to work they need to have a reference created
' to this template
' This template also required the Division.ini file whose path is also determined by the public constant in the template module


Public Sub CancelBttn_Click()
Unload UserForm1

End Sub



Public Sub OKBttn_Click()
' Check compulsory fields
txtUser = Trim(txtUser)
If txtUser = "" Then
    Beep
    MsgBox "This is a required field", vbOKOnly, "User Name"
    txtUser.SetFocus
    Exit Sub
End If

txtTitle = Trim(txtTitle)
If txtTitle = "" Then
    Beep
    MsgBox "This is a required field", vbOKOnly, "User's Title"
    txtTitle.SetFocus
    Exit Sub
End If
    
txtDDI = Trim(txtDDI)
If txtDDI = "" Then
    Beep
    MsgBox "This is a required field", vbOKOnly, "DDI Phone Number"
    txtDDI.SetFocus
    Exit Sub
End If

txtFax = Trim(txtFax)
If txtFax = "" Then
    Beep
    MsgBox "This is a required field", vbOKOnly, "Facsimile Number"
    txtFax.SetFocus
    Exit Sub
End If

If obPorirua.Value = False And obPataka.Value = False Then
    Beep
    MsgBox "Please choose a logo for signature."
    Exit Sub
End If

' Required to pick Web Address

'If cboWeb1.Value = "" Then
'    Beep
'    MsgBox "This is a required field", vbOKOnly, "Web address"
'    cboWeb1.SetFocus
'    Exit Sub
'End If
'
'If cboWeb2.Value = "" Then
'    Beep
'    MsgBox "This is a required field", vbOKOnly, "Web address"
'    cboWeb2.SetFocus
'    Exit Sub
'End If

' write the main user information to the ini file

System.PrivateProfileString(strDefaultUserIni, "UserSetup", "User") = txtUser
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Title") = txtTitle
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "DDI") = txtDDI
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Mobile") = txtMobile
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "BusGroup") = cboBusGroup
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Fax") = txtFax
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Email") = txtEmail
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Quote") = cboQuote
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Web1") = cboWeb1
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Link1") = IniOP.ReadFromINI("Links", cboWeb1.Value, strWebIni)
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Web2") = cboWeb2
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Link2") = IniOP.ReadFromINI("Links", cboWeb2.Value, strWebIni)
System.PrivateProfileString(strDefaultUserIni, "UserSetup", "DateCheck") = Format(Now(), "d mmmm yyyy")

' write UserA information to the ini file

System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "UserA") = txtUserA
System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "TitleA") = txtTitleA
System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "DDIA") = txtDDIA
System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "MobileA") = txtMobileA
System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "BusGroupA") = cboBusGroupA
System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "FaxA") = txtFaxA
System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "EmailA") = txtEmailA
System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "DateCheckA") = Format(Now(), "d mmmm yyyy")

' write UserB information to the ini file

System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "UserB") = txtUserB
System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "TitleB") = txtTitleB
System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "DDIB") = txtDDIB
System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "MobileB") = txtMobileB
System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "BusGroupB") = cboBusGroupB
System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "FaxB") = txtFaxB
System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "EmailB") = txtEmailB
System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "DateCheckB") = Format(Now(), "d mmmm yyyy")

' write UserC information to the ini file

System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "UserC") = txtUserC
System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "TitleC") = txtTitleC
System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "DDIC") = txtDDIC
System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "MobileC") = txtMobileC
System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "BusGroupC") = cboBusGroupC
System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "FaxC") = txtFaxC
System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "EmailC") = txtEmailC
System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "DateCheckC") = Format(Now(), "d mmmm yyyy")

' write UserD information to the ini file

System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "UserD") = txtUserD
System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "TitleD") = txtTitleD
System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "DDID") = txtDDID
System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "MobileD") = txtMobileD
System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "BusGroupD") = cboBusGroupD
System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "FaxD") = txtFaxD
System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "EmailD") = txtEmailD
System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "DateCheckD") = Format(Now(), "d mmmm yyyy")

' write UserE information to the ini file

System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "UserE") = txtUserE
System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "TitleE") = txtTitleE
System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "DDIE") = txtDDIE
System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "MobileE") = txtMobileE
System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "BusGroupE") = cboBusGroupE
System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "FaxE") = txtFaxE
System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "EmailE") = txtEmailE
System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "DateCheckE") = Format(Now(), "d mmmm yyyy")

' write UserF information to the ini file

System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "UserF") = txtUserF
System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "TitleF") = txtTitlef
System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "DDIF") = txtDDIF
System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "MobileF") = txtMobileF
System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "BusGroupF") = cboBusGroupF
System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "FaxF") = txtFaxF
System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "EmailF") = txtEmailF
System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "DateCheckF") = Format(Now(), "d mmmm yyyy")

System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Userg") = txtUserG
System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Titleg") = txtTitleG
System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "DDIg") = txtDDIG
System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Mobileg") = txtMobileG
System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "BusGroupg") = cboBusGroupG
System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Faxg") = txtFaxG
System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Emailg") = txtEmailG
System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "DateCheckg") = Format(Now(), "d mmmm yyyy")

System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "Userh") = txtUserH
System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "Titleh") = txtTitleh
System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "DDIh") = txtDDIH
System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "Mobileh") = txtMobileH
System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "BusGrouph") = cboBusGroupH
System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "Faxh") = txtFaxH
System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "Emailh") = txtEmailH
System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "DateCheckh") = Format(Now(), "d mmmm yyyy")

System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "Useri") = txtUserI
System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "Titlei") = txtTitleI
System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "DDIi") = txtDDII
System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "Mobilei") = txtMobileI
System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "BusGroupi") = cboBusGroupI
System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "Faxi") = txtFaxI
System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "Emaili") = txtEmailI
System.PrivateProfileString(strDefaultUserIni, "UserSetupi", "DateChecki") = Format(Now(), "d mmmm yyyy")



Unload UserForm1
'PCC Setting
'Run script to generate signature
'Shell "Wscript.exe W:\!Common\Templates\Allfields_Setup\2010_Signature\PCC_Signature.vbs"
'new approach based on VBS script. 17/08/2017 tao@allfields.co.nz
   
    Dim doc As Document
    Dim sSig As String  'builing block name
    Set doc = Documents.Add(, , , True)
    Set oEmail = Application.EmailOptions
    Set oSignature = oEmail.EmailSignature
    Set oSignatureEntry = oSignature.EmailSignatureEntries
    
    Dim rg As Range
    'ThisDocument.Content.Delete
    Set rg = doc.Content
    rg.Collapse wdCollapseStart
    sSig = IIf(obPorirua, "signature", "signature2")
    ThisDocument.AttachedTemplate.BuildingBlockEntries(sSig).Insert rg, True
    Dim names() As String
    names = Split(txtUser)
    FillBookmark doc, "firstname", names(0)
    If UBound(names) > 0 Then
        FillBookmark doc, "lastname", names(1)
    Else
        FillBookmark doc, "lastname", ""
    End If
    
    FillBookmark doc, "ddi", txtDDI
    FillBookmark doc, "mobile", txtMobile
    FillBookmark doc, "email", txtEmail
    FillBookmark doc, "title", txtTitle
    For Each para In doc.Paragraphs
        para.SpaceAfter = 0
    Next para
    oSignatureEntry.Add "PCC-Test", doc.Content
    oSignature.NewMessageSignature = "PCC-Test"
    oSignature.ReplyMessageSignature = "PCC-Test"
    doc.Saved = True
    doc.Close
    Set doc = Nothing
'Allfields Test
'Run script to generate signature
'Shell "Wscript.exe H:\PCC_Signature.vbs"

MsgBox "You have just now set your PCC Email Signature,  Check it out!  If you need to amend your information at anytime please access the Profile Setup Button on the Council Templates tab on the Ribbon", vbInformation, "Update Your Information?"


End Sub
' Help for User
Public Sub txtUser_Enter()
txtHelp.Caption = "Please enter the usual name for use in letters, your email signature and other correspondence." + vbCr + "For example 'John Cooper'"
End Sub
Public Sub txtTitle_Enter()
txtHelp.Caption = "Please enter in the title to be used in general correspondence." + vbCr + "For example 'Project Manager'"
End Sub
Public Sub txtDDI_Enter()
txtHelp.Caption = "Please enter your full DDI phone number." + vbCr + vbCr + "Please use the format 64-4-474 3000."
End Sub
Public Sub txtFax_Enter()
txtHelp.Caption = "Please enter your facsimile phone number." + vbCr + vbCr + "Please use the format 64-4-474 3035."
End Sub
Public Sub cboBusGroup_Enter()
txtHelp.Caption = "Please select a business group from the list." + vbCr + vbCr + "Please choose 'None' if none is suitable."
End Sub
Public Sub txtMobile_Enter()
txtHelp.Caption = "Optionally, please enter in your Mobile phone number."
End Sub
Public Sub txtEmail_Enter()
txtHelp.Caption = "Please enter in your email address."
End Sub
Public Sub cboquote_Enter()
txtHelp.Caption = "This quote will be displayed at the bottom of you email signature." + vbCr + vbCr + "You can come back here and pick another quote at any time."
End Sub
' Help for UserA
Public Sub txtUserA_Enter()
txtHelp.Caption = "Please enter the usual name for use in letters and other correspondence." + vbCr + "For example 'John Cooper'"
End Sub
Public Sub txtTitleA_Enter()
txtHelp.Caption = "Please enter in the title to be used in general correspondence." + vbCr + "For example 'Project Manager'"
End Sub
Public Sub txtDDIA_Enter()
txtHelp.Caption = "Please enter your full DDI phone number." + vbCr + vbCr + "Please use the format 64-4-474 3000."
End Sub
Public Sub txtFaxA_Enter()
txtHelp.Caption = "Please enter your facsimile phone number." + vbCr + vbCr + "Please use the format 64-4-474 3035."
End Sub
Public Sub cboBusGroupA_Enter()
txtHelp.Caption = "Please select a business group from the list." + vbCr + vbCr + "Please choose 'None' if none is suitable."
End Sub
Public Sub txtMobileA_Enter()
txtHelp.Caption = "Optionally, please enter in your Mobile phone number."
End Sub
Public Sub txtEmailA_Enter()
txtHelp.Caption = "Please enter in your email address."
End Sub
' Help for UserB
Public Sub txtUserB_Enter()
txtHelp.Caption = "Please enter the usual name for use in letters and other correspondence." + vbCr + "For example 'John Cooper'"
End Sub
Public Sub txtTitleB_Enter()
txtHelp.Caption = "Please enter in the title to be used in general correspondence." + vbCr + "For example 'Project Manager'"
End Sub
Public Sub txtDDIB_Enter()
txtHelp.Caption = "Please enter your full DDI phone number." + vbCr + vbCr + "Please use the format 64-4-474 3000."
End Sub
Public Sub txtFaxB_Enter()
txtHelp.Caption = "Please enter your facsimile phone number." + vbCr + vbCr + "Please use the format 64-4-474 3035."
End Sub
Public Sub cboBusGroupB_Enter()
txtHelp.Caption = "Please select a business group from the list." + vbCr + vbCr + "Please choose 'None' if none is suitable."
End Sub
Public Sub txtMobileB_Enter()
txtHelp.Caption = "Optionally, please enter in your Mobile phone number."
End Sub
Public Sub txtEmailB_Enter()
txtHelp.Caption = "Please enter in your email address."
End Sub
' Help for UserC
Public Sub txtUserC_Enter()
txtHelp.Caption = "Please enter the usual name for use in letters and other correspondence." + vbCr + "For example 'John Cooper'"
End Sub
Public Sub txtTitleC_Enter()
txtHelp.Caption = "Please enter in the title to be used in general correspondence." + vbCr + "For example 'Project Manager'"
End Sub
Public Sub txtDDIC_Enter()
txtHelp.Caption = "Please enter your full DDI phone number." + vbCr + vbCr + "Please use the format 64-4-474 3000."
End Sub
Public Sub txtFaxC_Enter()
txtHelp.Caption = "Please enter your facsimile phone number." + vbCr + vbCr + "Please use the format 64-4-474 3035."
End Sub
Public Sub cboBusGroupC_Enter()
txtHelp.Caption = "Please select a business group from the list." + vbCr + vbCr + "Please choose 'None' if none is suitable."
End Sub
Public Sub txtMobileC_Enter()
txtHelp.Caption = "Optionally, please enter in your Mobile phone number."
End Sub
Public Sub txtEmailC_Enter()
txtHelp.Caption = "Please enter in your email address."
End Sub
' Help for UserD
Public Sub txtUserD_Enter()
txtHelp.Caption = "Please enter the usual name for use in letters and other correspondence." + vbCr + "For example 'John Cooper'"
End Sub
Public Sub txtTitleD_Enter()
txtHelp.Caption = "Please enter in the title to be used in general correspondence." + vbCr + "For example 'Project Manager'"
End Sub
Public Sub txtDDID_Enter()
txtHelp.Caption = "Please enter your full DDI phone number." + vbCr + vbCr + "Please use the format 64-4-474 3000."
End Sub
Public Sub txtFaxD_Enter()
txtHelp.Caption = "Please enter your facsimile phone number." + vbCr + vbCr + "Please use the format 64-4-474 3035."
End Sub
Public Sub cboBusGroupD_Enter()
txtHelp.Caption = "Please select a business group from the list." + vbCr + vbCr + "Please choose 'None' if none is suitable."
End Sub
Public Sub txtMobileD_Enter()
txtHelp.Caption = "Optionally, please enter in your Mobile phone number."
End Sub
Public Sub txtEmailD_Enter()
txtHelp.Caption = "Please enter in your email address."
End Sub
' Help for UserE
Public Sub txtUserE_Enter()
txtHelp.Caption = "Please enter the usual name for use in letters and other correspondence." + vbCr + "For example 'John Cooper'"
End Sub
Public Sub txtTitleE_Enter()
txtHelp.Caption = "Please enter in the title to be used in general correspondence." + vbCr + "For example 'Project Manager'"
End Sub
Public Sub txtDDIE_Enter()
txtHelp.Caption = "Please enter your full DDI phone number." + vbCr + vbCr + "Please use the format 64-4-474 3000."
End Sub
Public Sub txtFaxE_Enter()
txtHelp.Caption = "Please enter your facsimile phone number." + vbCr + vbCr + "Please use the format 64-4-474 3035."
End Sub
Public Sub cboBusGroupE_Enter()
txtHelp.Caption = "Please select a business group from the list." + vbCr + vbCr + "Please choose 'None' if none is suitable."
End Sub
Public Sub txtMobileE_Enter()
txtHelp.Caption = "Optionally, please enter in your Mobile phone number."
End Sub
Public Sub txtEmailE_Enter()
txtHelp.Caption = "Please enter in your email address."
End Sub
' Help for UserF
Public Sub txtUserF_Enter()
txtHelp.Caption = "Please enter the usual name for use in letters and other correspondence." + vbCr + "For example 'John Cooper'"
End Sub
Public Sub txtTitleF_Enter()
txtHelp.Caption = "Please enter in the title to be used in general correspondence." + vbCr + "For example 'Project Manager'"
End Sub
Public Sub txtDDIF_Enter()
txtHelp.Caption = "Please enter your full DDI phone number." + vbCr + vbCr + "Please use the format 64-4-474 3000."
End Sub
Public Sub txtFaxF_Enter()
txtHelp.Caption = "Please enter your facsimile phone number." + vbCr + vbCr + "Please use the format 64-4-474 3035."
End Sub
Public Sub cboBusGroupF_Enter()
txtHelp.Caption = "Please select a business group from the list." + vbCr + vbCr + "Please choose 'None' if none is suitable."
End Sub
Public Sub txtMobileF_Enter()
txtHelp.Caption = "Optionally, please enter in your Mobile phone number."
End Sub
Public Sub txtEmailF_Enter()
txtHelp.Caption = "Please enter in your email address."
End Sub

Public Sub OKBttn_Enter()
txtHelp.Caption = ""
End Sub
Public Sub CancelBttn_Enter()
txtHelp.Caption = ""
End Sub

Public Sub UserForm_Initialize()

FilePaths.Autoexec
' Populate main part of form


If txtUser = "" Then
    txtUser = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "User")
End If
If txtTitle = "" Then
    txtTitle = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Title")
End If
If txtDDI = "" Then
    txtDDI = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "DDI")
End If
If txtFax = "" Then
    txtFax = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Fax")
End If
If txtMobile = "" Then
    txtMobile = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Mobile")
End If
If txtEmail = "" Then
    txtEmail = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Email")
End If

'If cboQuote = "" Then
  '  cboQuote = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Quote")
'End If

' Populate UserA

If txtUserA = "" Then
    txtUserA = System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "UserA")
End If
If txtTitleA = "" Then
    txtTitleA = System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "TitleA")
End If
If txtDDIA = "" Then
    txtDDIA = System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "DDIA")
End If
If txtFaxA = "" Then
    txtFaxA = System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "FaxA")
End If
If txtMobileA = "" Then
    txtMobileA = System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "MobileA")
End If
If txtEmailA = "" Then
    txtEmailA = System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "EmailA")
End If

' Populate UserB

If txtUserB = "" Then
    txtUserB = System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "UserB")
End If
If txtTitleB = "" Then
    txtTitleB = System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "TitleB")
End If
If txtDDIB = "" Then
    txtDDIB = System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "DDIB")
End If
If txtFaxB = "" Then
    txtFaxB = System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "FaxB")
End If
If txtMobileB = "" Then
    txtMobileB = System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "MobileB")
End If
If txtEmailB = "" Then
    txtEmailB = System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "EmailB")
End If

' Populate UserC

If txtUserC = "" Then
    txtUserC = System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "UserC")
End If
If txtTitleC = "" Then
    txtTitleC = System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "TitleC")
End If
If txtDDIC = "" Then
    txtDDIC = System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "DDIC")
End If
If txtFaxC = "" Then
    txtFaxC = System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "FaxC")
End If
If txtMobileC = "" Then
    txtMobileC = System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "MobileC")
End If
If txtEmailC = "" Then
    txtEmailC = System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "EmailC")
End If


' Populate UserD

If txtUserD = "" Then
    txtUserD = System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "UserD")
End If
If txtTitleD = "" Then
    txtTitleD = System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "TitleD")
End If
If txtDDID = "" Then
    txtDDID = System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "DDID")
End If
If txtFaxD = "" Then
    txtFaxD = System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "FaxD")
End If
If txtMobileD = "" Then
    txtMobileD = System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "MobileD")
End If
If txtEmailD = "" Then
    txtEmailD = System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "EmailD")
End If

' Populate UserE

If txtUserE = "" Then
    txtUserE = System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "UserE")
End If
If txtTitleE = "" Then
    txtTitleE = System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "TitleE")
End If
If txtDDIE = "" Then
    txtDDIE = System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "DDIE")
End If
If txtFaxE = "" Then
    txtFaxE = System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "FaxE")
End If
If txtMobileE = "" Then
    txtMobileE = System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "MobileE")
End If
If txtEmailE = "" Then
    txtEmailE = System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "EmailE")
End If

' Populate UserF

If txtUserF = "" Then
    txtUserF = System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "UserF")
End If
If txtTitlef = "" Then
    txtTitlef = System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "TitleF")
End If
If txtDDIF = "" Then
    txtDDIF = System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "DDIF")
End If
If txtFaxF = "" Then
    txtFaxF = System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "FaxF")
End If
If txtMobileF = "" Then
    txtMobileF = System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "MobileF")
End If
If txtEmailF = "" Then
    txtEmailF = System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "EmailF")
End If

' Populate Userg

If txtUserG = "" Then
    txtUserG = System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Userg")
End If
If txtTitleG = "" Then
    txtTitleG = System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Titleg")
End If
If txtDDIG = "" Then
    txtDDIG = System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "DDIg")
End If
If txtFaxG = "" Then
    txtFaxG = System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Faxg")
End If
If txtMobileG = "" Then
    txtMobileG = System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Mobileg")
End If
If txtEmailG = "" Then
    txtEmailG = System.PrivateProfileString(strDefaultUserIni, "UserSetupg", "Emailg")
End If

' Populate UserH

If txtUserH = "" Then
    txtUserH = System.PrivateProfileString(strDefaultUserIni, "UserSetuph", "Userh")
End If
If txtTitleh = "" Then
    txtTitleh = System.PrivateProfileString(strDefaultUserIni, "UserSetupH", "TitleH")
End If
If txtDDIH = "" Then
    txtDDIH = System.PrivateProfileString(strDefaultUserIni, "UserSetupH", "DDIH")
End If
If txtFaxH = "" Then
    txtFaxH = System.PrivateProfileString(strDefaultUserIni, "UserSetupH", "FaxH")
End If
If txtMobileH = "" Then
    txtMobileH = System.PrivateProfileString(strDefaultUserIni, "UserSetupH", "MobileH")
End If
If txtEmailH = "" Then
    txtEmailH = System.PrivateProfileString(strDefaultUserIni, "UserSetupH", "EmailH")
End If

' Populate UserI

If txtUserI = "" Then
    txtUserI = System.PrivateProfileString(strDefaultUserIni, "UserSetupI", "UserI")
End If
If txtTitleI = "" Then
    txtTitleI = System.PrivateProfileString(strDefaultUserIni, "UserSetupI", "TitleI")
End If
If txtDDII = "" Then
    txtDDII = System.PrivateProfileString(strDefaultUserIni, "UserSetupI", "DDII")
End If
If txtFaxI = "" Then
    txtFaxI = System.PrivateProfileString(strDefaultUserIni, "UserSetupI", "FaxI")
End If
If txtMobileI = "" Then
    txtMobileI = System.PrivateProfileString(strDefaultUserIni, "UserSetupI", "MobileI")
End If
If txtEmailI = "" Then
    txtEmailI = System.PrivateProfileString(strDefaultUserIni, "UserSetupI", "EmailI")
End If


'--------------------------------------------
'populate quote choice
Dim clsFileOp2 As New FileOperations

' Dim Control and Populate It
Dim ctlBusGroup2 As control
Set ctlBusGroup2 = Me.cboQuote

lret = clsFileOp2.PopulateCtl(strQuoteIni, ctlBusGroup2, "[Quote]")
If lret = False Then
    MsgBox ("Could not open " & strQuoteIni)
End If
''''''''end quote ini population
'----------------------------------------------


'populate business groups

Dim clsFileOp As New FileOperations

' Dim Control and Populate It
Dim ctlBusGroup As control
Set ctlBusGroup = Me.cboBusGroup

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

' Populate other combo boxes

Set ctlBusGroup = Me.cboBusGroupA

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupB

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupC

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupD

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupE

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupF

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupG

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupH

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

Set ctlBusGroup = Me.cboBusGroupI

lret = clsFileOp.PopulateCtl(strDivisionIni, ctlBusGroup, "[Division]")
If lret = False Then
    MsgBox ("Could not open " & strDivisionIni)
End If

' Read from ini file the current setting of the business group

cboBusGroup = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "BusGroup")
cboBusGroupA = System.PrivateProfileString(strDefaultUserIni, "UserSetupA", "BusGroupA")
cboBusGroupB = System.PrivateProfileString(strDefaultUserIni, "UserSetupB", "BusGroupB")
cboBusGroupC = System.PrivateProfileString(strDefaultUserIni, "UserSetupC", "BusGroupC")
cboBusGroupD = System.PrivateProfileString(strDefaultUserIni, "UserSetupD", "BusGroupD")
cboBusGroupE = System.PrivateProfileString(strDefaultUserIni, "UserSetupE", "BusGroupE")
cboBusGroupF = System.PrivateProfileString(strDefaultUserIni, "UserSetupF", "BusGroupF")

' Read from ini file the current setting of the business group
cboQuote = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Quote")


'Populate Web Addresses from WebIni
Dim WebNames() As String
Dim i As Integer
Dim web1Val As String
Dim web2Val As String
web1Val = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Web1")
web2Val = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "Web2")
WebNames = IniOP.LoadIniSectionKeysArray("Links", strWebIni)
For i = LBound(WebNames) To UBound(WebNames)
    cboWeb1.AddItem WebNames(i), i
    cboWeb2.AddItem WebNames(i), i
    If web1Val = WebNames(i) Then cboWeb1.ListIndex = i
    If web2Val = WebNames(i) Then cboWeb2.ListIndex = i
Next i
'' to here

MultiPage1.Value = 0
txtUser.SetFocus
End Sub


Public Sub ReplaceTextInSelection(sFind As String, sReplace As String)
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = sFind
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = sFind
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub


