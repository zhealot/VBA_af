VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProfileInformation 
   Caption         =   "Profile Setup"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   OleObjectBlob   =   "MVCOT_startup_frmProfileInformation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProfileInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const REQUIRED_COLOUR = 8421631
Const ADD_NEW_LABEL = "Add New..."
Dim aFields

Public Sub SelectUser(strUser As String)
    Me.cboUserPickList.Value = strUser
End Sub

Private Sub UserForm_Initialize()
    FilePaths.Autoexec
    aFields = Array("Name", "Title", "DDI", "Fax", _
            "Email", "Afterhours", "BusinessGroup", _
            "BusinessStreetAddress", "BusinessPostalAddress", _
            "BusinessCity", "BusinessPhone", "BusinessFacsimile")
    
    ' Check the Old UserINI does not exist and import if needed
    PopulateWorkUnits cboWorkUnit
    UpdateUserPicker
End Sub

Sub UpdateUserPicker(Optional strSelected As String = "")
    Dim aUsers As Variant
    Dim iPos As Integer
    aUsers = IniOP.LoadIniSectionsArray(strUserIni)
    
    Me.cboUserPickList.Clear
    For iPos = LBound(aUsers) To UBound(aUsers)
        Me.cboUserPickList.AddItem aUsers(iPos)
    Next iPos
    Me.cboUserPickList.AddItem ADD_NEW_LABEL
    
    If strSelected <> "" Then
        Me.cboUserPickList.Value = strSelected
    ElseIf Me.cboUserPickList.ListCount > 0 Then
        Me.cboUserPickList.ListIndex = 0
    End If
End Sub

Sub PopulateUserForm()
    Dim strUser As String
    Dim iFPos As Integer
    Dim strField As String
    strUser = Me.cboUserPickList.Value
    
    Me.cboWorkUnit.ListIndex = 0
    If strUser <> ADD_NEW_LABEL Then _
        Me.cboWorkUnit.Value = GetUserSetting("Workunit", strUser)

    For iFPos = LBound(aFields) To UBound(aFields)
        strField = CStr(aFields(iFPos))
        Debug.Print "txt" & strField & " = " & GetUserSetting(strField, strUser)
        If strUser <> ADD_NEW_LABEL Then
            Me.Controls("txt" & strField).Value = GetUserSetting(strField, strUser)
        Else
            Me.Controls("txt" & strField).Value = ""
        End If
    Next iFPos
    
    CheckAddress
    
    Me.cmdSave.Enabled = False
    
End Sub

'*
'  Check the Address has been filled in, if not
'  based on the selected workunit see if it can be
'  pre-populated.
'*
Sub CheckAddress()
    If Me.txtBusinessStreetAddress = "" And Me.txtBusinessPostalAddress = "" Then
        FillAddressFromGlobal
    End If
End Sub

Sub FillAddressFromGlobal()
    Dim strStreet As String, strPostal As String, strCity As String
    Dim strPhone As String, strFax As String, strGroup As String
    strGroup = Me.cboWorkUnit.Value

    Me.txtBusinessStreetAddress = GetINISetting(strAddressConfigPath, strGroup, "Street")
    Me.txtBusinessPostalAddress = GetINISetting(strAddressConfigPath, strGroup, "Postal")
    ' there is no special setting for this WU, use default
    If Me.txtBusinessStreetAddress = "" And Me.txtBusinessPostalAddress = "" Then
        strGroup = "Default"
        Me.txtBusinessStreetAddress = GetINISetting(strAddressConfigPath, strGroup, "Street")
        Me.txtBusinessPostalAddress = GetINISetting(strAddressConfigPath, strGroup, "Postal")
    End If
    
    Me.txtBusinessCity = GetINISetting(strAddressConfigPath, strGroup, "City")
    Me.txtBusinessPhone = GetINISetting(strAddressConfigPath, strGroup, "Telephone")
    Me.txtBusinessFacsimile = GetINISetting(strAddressConfigPath, strGroup, "Facsimile")
End Sub

Private Sub cmdSave_Click()
    If Not CheckMandatoryFields(Me) Then Exit Sub
        
    Dim strUser As String
    Dim strField As String
    Dim iFPos As Integer
    
    strUser = Me.cboUserPickList.Value
    If strUser = ADD_NEW_LABEL Then strUser = Me.txtName
    
    For iFPos = LBound(aFields) To UBound(aFields)
        strField = CStr(aFields(iFPos))
        WriteUserSetting strField, Me.Controls("txt" & strField).Value, strUser
    Next iFPos
    WriteUserSetting "Workunit", Me.cboWorkUnit.Value, strUser
    
    If strUser <> Me.txtName Then _
        FilePaths.IniOP.RenameIniSection strUser, Me.txtName, strUserIni
    strUser = Me.txtName
    
    UpdateUserPicker (strUser)
    
End Sub

Private Sub cboUserPickList_Change()
    Me.cmdDelete.Enabled = Me.cboUserPickList.Value <> ADD_NEW_LABEL
    PopulateUserForm
End Sub

Private Sub cboWorkUnit_Change()
    FillAddressFromGlobal
End Sub

Private Sub cmdDelete_Click()
    Dim delete
    
    delete = MsgBox("Are you sure you want to delete " & Me.cboUserPickList.Value & "?" _
                & vbCr & vbCr & "This can not be undone.", vbYesNo, "Are you sure?")
    
    If delete = vbYes Then
        IniOP.DeleteIniKey "UserIndex", Me.cboUserPickList, strUserIni
        IniOP.DeleteIniSection Me.cboUserPickList.Value, strUserIni
    End If
    
    UpdateUserPicker
    
End Sub
Private Sub txtBusinessStreetAddress_Change()
    If txtBusinessStreetAddress.Value <> "" Then
        txtBusinessPostalAddress.Object.BorderColor = STANDARD_COLOUR
    Else
        txtBusinessPostalAddress.Object.BorderColor = REQUIRED_COLOUR
    End If
End Sub

Private Sub txtBusinessPostalAddress_Change()
    If txtBusinessPostalAddress.Value <> "" Then
        txtBusinessStreetAddress.Object.BorderColor = STANDARD_COLOUR
    Else
        txtBusinessStreetAddress.Object.BorderColor = REQUIRED_COLOUR
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ProfileInfo_Enter()
    EnableSave
End Sub

Private Sub fraAddress_Enter()
    EnableSave
End Sub

Private Sub fraBusiness_Enter()
    EnableSave
End Sub

Sub EnableSave()
    Me.cmdSave.Enabled = True
End Sub
