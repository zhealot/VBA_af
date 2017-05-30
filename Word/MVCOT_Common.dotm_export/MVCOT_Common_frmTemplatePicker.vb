VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplatePicker 
   Caption         =   "MVCOT Templates"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   OleObjectBlob   =   "MVCOT_Common_frmTemplatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------
' These templates have been prepared and developed for the MED
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     info@allfields.co.nz, 04 978 7101
' Date:             March 2011
' Description:      Form used for picking template to load. Scans the
'                   Workgroup Templates folder for templates with a .dotm
'                   extension, groups them by the text before the dash, and
'                   named by the text after the dash, less the extension
'-----------------------------------------------------------------------------
Option Explicit

Dim MenuCancelled As Boolean, strCurrentUserFolder As String
Dim optControls() As New OptionButtonHandler

Private Sub cboUserPickList_Change()
    Me.txtProfileWorkUnit.Value = GetUserSetting("Workunit", cboUserPickList.Value)
End Sub

Private Sub EditProfileSettings_Click()
    Dim SelectedUser As String
    SelectedUser = Me.cboUserPickList.Value
    LaunchProfileSetup SelectedUser
    ' afterwards refresh the picker in-case the name has been changed
    LoadUserList
End Sub


'------------------------------------------------------------
'Userform Initialize
'When the userform first displays, parse the workgroup templates
'
'------------------------------------------------------------
Private Sub UserForm_Initialize()
    Const TEMPLATE_EXTENSION = "dotm"
    Const EXTENSION_LENGTH = 4
    Const LEFT_PADDING = 10
    Const GRID_LEFT = 20
    Const GRID_TOP = 90
    
    Dim GRID_COLS As Integer
    GRID_COLS = 2
    Dim GROUP_WIDTH As Integer
    GROUP_WIDTH = 100
    Dim ITEM_WIDTH As Integer
    ITEM_WIDTH = 95
    Dim ITEM_HEIGHT As Integer
    ITEM_HEIGHT = 15
    
    Dim strCurrentGroup As String
    Dim strGroup As String
    Dim strTemplate As String
    Dim strTemplateFile As String
    Dim strExtensionless As String
    Dim optCurrent As MSForms.OptionButton
    Dim lblGroup As MSForms.label
    Dim iDashPos As Integer
    Dim iLeft As Integer, iTop As Integer
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim iIndex As Integer
    Dim iGroupCount As Integer
    Dim iRowMax As Integer
    Dim iRowTop As Integer

    Dim i As Integer

    Dim allFiles As Variant
    allFiles = GetFileList(strCoreTemplatesPath & "\*." & TEMPLATE_EXTENSION)
    If IsArray(allFiles) Then
        Call QuickSort(allFiles, LBound(allFiles), UBound(allFiles))
    Else
        ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
            & vbCr & vbCr & "Looking in " & """" & strCoreTemplatesPath & """"
    End If
    
    iColumn = 0
    iRow = 0
    iLeft = (-GROUP_WIDTH) + GRID_LEFT
    iTop = GRID_TOP - ITEM_HEIGHT
    iRowTop = GRID_TOP
    iIndex = 0
    iGroupCount = 0
    iRowMax = 0
    ' Start the loop.
    For i = 1 To UBound(allFiles)
        
        ' Set the template and position
        iTop = iTop + ITEM_HEIGHT
        iIndex = iIndex + 1
        strTemplateFile = allFiles(i)
        iRowMax = MAX(iRowMax, iTop)
        
        ' Remove the extension
        iDashPos = InStr(1, strTemplateFile, "-")
        If iDashPos = 0 Then ThrowFatalError _
            "Error loading templates. Template filenames must contain a hyphen" & vbCr & _
            "Loading " & strTemplateFile
        strExtensionless = Left(strTemplateFile, Len(strTemplateFile) - (EXTENSION_LENGTH + 1))
        strGroup = Left(strExtensionless, (iDashPos - 1))
        strTemplate = Right(strExtensionless, Len(strExtensionless) - (iDashPos + 1))
               
        If strGroup <> strCurrentGroup Then
            If iColumn > GRID_COLS Then
                iLeft = (-GROUP_WIDTH) + GRID_LEFT
                iColumn = 0
                iRow = iRow + 1
                iRowTop = iRowMax + 20
                iRowMax = 0
            End If
            iColumn = iColumn + 1
            iLeft = iLeft + (GROUP_WIDTH)
            strCurrentGroup = strGroup
            iTop = iRowTop
            Set lblGroup = frmTemplatePicker.Controls.Add("Forms.Label.1", "lbl" & iIndex, True)
            With lblGroup
                .Caption = strGroup
                .Top = iTop
                .Left = iLeft
                .Width = ITEM_WIDTH
                .Font.Bold = True
                .Font.Size = 10
                '.BackColor = RGB(255, 0, 0)
            End With
            iTop = iTop + ITEM_HEIGHT
            iGroupCount = iGroupCount + 1
        End If

        Set optCurrent = frmTemplatePicker.Controls.Add("Forms.OptionButton.1", "opt" & iIndex, True)
        With optCurrent
            .Caption = strTemplate
            .Top = iTop
            .Left = iLeft
            .Width = ITEM_WIDTH
            .GroupName = "ga"
            .Tag = strTemplateFile
        End With
        ReDim Preserve optControls(1 To iIndex)
        Set optControls(iIndex).cb = optCurrent
        optControls(iIndex).Caption = strExtensionless
        Debug.Print optCurrent.Caption
    Next i
    
    If iIndex < 1 Then
        ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
            & vbCr & vbCr & "Looking in " & """" & strCoreTemplatesPath & """"
    End If
    
    iTop = iTop + ITEM_HEIGHT
    iRowMax = MAX(iRowMax, iTop)
    
    ' Set the form boundaries to fit the contents
'    Me.Width = (GRID_LEFT * 2) + ((GRID_COLS + 1) * GROUP_WIDTH)
'    If Me.Width < 450 Then Me.Width = 450
'
'    If iGroupCount Mod (GRID_COLS + 1) <> 0 Then
'        Me.Height = (iRowMax) + 50
'    Else
'        Me.Height = (iRowMax) + 100
'    End If
'    buttonFrame.Top = Me.Height - (buttonFrame.Height + 25)
'    buttonFrame.Left = Me.Width - buttonFrame.Width - 10
    
    ' Set the default user information
    FormHelper.PopulateUserPicker cboUserPickList
    
End Sub

Sub cmbCancel_Click()
    Unload Me
    End
End Sub

Sub cmbOK_Click()
    Dim optControl As control
    Dim strTemplateChoice As String
    Dim strFileToCreate As String
    Dim boolFound As Boolean
    Dim strWorkunit As String
    
    boolFound = False
    For Each optControl In Me.Controls
        If (Left(optControl.Name, 3) = "opt") Then
            If optControl.Value = True Then
                strTemplateChoice = optControl.Tag
                boolFound = True
                Exit For
            End If
        End If
    Next optControl
    
    If Not boolFound Then
        MsgBox "Please select a template from the list", vbOKOnly + vbCritical, "No template selected"
        Exit Sub
    End If
    
    strWorkunit = GetUserSetting("Workunit", Me.cboUserPickList)
    If FileExists(strSpecialTemplatesPath & "\" & strWorkunit & "\" & strTemplateChoice) Then
        strFileToCreate = strSpecialTemplatesPath & "\" & strWorkunit & "\" & strTemplateChoice
    Else
        strFileToCreate = strCoreTemplatesPath & "\" & strTemplateChoice
    End If
    Unload Me
        
    SetActiveUser Me.cboUserPickList.Value
    
    'Create new document
    Documents.Add Template:=strFileToCreate

End Sub

Function MAX(iNum1 As Integer, iNum2 As Integer) As Integer
    If iNum1 > iNum2 Then
        MAX = iNum1
    Else
        MAX = iNum2
    End If
End Function

Sub LoadUserList()
    PopulateUserPicker Me.cboUserPickList
End Sub
