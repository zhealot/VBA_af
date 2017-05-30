Attribute VB_Name = "Ribbon"
' **************************************************
'
' Ribbon Control Module for MSD template toolkit.
' Responsible for delegating button-presses on the ribbon to
' the respective macros.
'
' @author Patrick, Allfields Customised Solutions, 04 978 7101
'
' ***************************************************


Public Sub ShowTemplatesMenu(control As IRibbonControl)
    FilePaths.Autoexec
    CheckRequirements
    LaunchTemplatePicker
End Sub
'
'Public Sub ShowOtherTemplatesMenu(control As IRibbonControl)
'    FilePaths.Autoexec
'    CheckRequirements
'    If Not FolderExists(strOtherTemplatesPath) Then
'        ThrowFatalError "Cannot load Other Templates directory. " & vbCr & _
'            "Looking in " & strWorkgroupTemplatesPath
'    End If
'    LaunchOtherTemplatesPicker
'End Sub
'
'Public Sub OpenCorporateTemplates(control As IRibbonControl)
'    OpenURL HYPERLINK_CORPORATE_TEMPLATES
'End Sub
'
Public Sub ShowProfileSetup(control As IRibbonControl)
    FilePaths.Autoexec
    If Not CheckRequirements Then _
        LaunchProfileSetup
End Sub
'
'Public Sub OpenStyleGuide(control As IRibbonControl)
'    FilePaths.Autoexec
'    CheckRequirements
'    OpenURL GetINISetting(strMSDGlobalPath, "Ribbon", "StyleGuideURL")
'End Sub

Public Sub OpenHowToGuide(control As IRibbonControl)
    FilePaths.Autoexec
    CheckRequirements
    'OpenURL GetINISetting(strMSDGlobalPath, "Ribbon", "HowToGuideURL")
    Documents.Add strHelpFile
End Sub

Public Sub Images(control As IRibbonControl)
    FilePaths.Autoexec
    If Dir(strImagePath, vbDirectory) = "" Or strImagePath = "" Then
        MsgBox "Image folder has not been found"
    Else
        Shell "explorer.exe" & " " & strImagePath, vbNormalFocus
    End If
End Sub

'Public Sub ShowVersionInformation(control As IRibbonControl)
'    Call MsgBox("Toolkit version " & FILE_VERSION & vbCr & vbCr _
'            & "Last updated " & LAST_UPDATED & vbCr _
'            & "by " & LAST_AUTHOR, vbInformation, "Template Toolkit Version Information")
'End Sub

' Utility subs so we can launch the forms without the ribbon args
Public Sub LaunchTemplatePicker()
    Load frmTemplatePicker
    frmTemplatePicker.Show
End Sub
Public Sub LaunchOtherTemplatesPicker()
    Load frmOtherTemplatesPicker
    frmOtherTemplatesPicker.Show
End Sub
Public Sub LaunchProfileSetup(Optional SelectedUser As String = "")
    Load frmProfileInformation
    If SelectedUser <> "_none" Then
        If SelectedUser = "" Then SelectedUser = Allfields.GetDefaultUser
        frmProfileInformation.SelectUser SelectedUser
    End If
    frmProfileInformation.Show
    
End Sub

Private Sub OpenURL(strURL As String)

    On Error GoTo Error_Handler
    ActiveDocument.FollowHyperlink strURL, , True
        
    Exit Sub
        
Error_Handler:
    ThrowFatalError ("Could not open " & strURL & vbCr & vbCr & "Please ensure you have an internet connection available")
  
End Sub
