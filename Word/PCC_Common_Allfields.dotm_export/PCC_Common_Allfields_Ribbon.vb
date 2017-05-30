Attribute VB_Name = "Ribbon"
' **************************************************
' Ribbon Control Module for PCC template toolkit.
' Responsible for delegating button-presses on the ribbon to
' the respective macros.
' @author Patrick, Allfields Customised Solutions, 04 978 7101
' ***************************************************.


Public Sub ShowTemplatesMenu(Control As IRibbonControl)
    FilePaths.Autoexec
    CheckRequirements
    LaunchTemplatePicker
End Sub

Public Sub ShowProfileSetup(Control As IRibbonControl)
    FilePaths.Autoexec
    If Not CheckRequirements(True) Then _
        LaunchProfileSetup
End Sub

Public Sub OpenHowToGuide(Control As IRibbonControl)
    Documents.Add Template:="Instructions.dotx"
End Sub

Public Sub ShowVersionInformation(Control As IRibbonControl)
    Call MsgBox("Toolkit version " & FILE_VERSION & vbCr & vbCr _
            & "Last updated " & LAST_UPDATED & vbCr _
            & "by " & LAST_AUTHOR, vbInformation, "Template Toolkit Version Information")
End Sub

Public Sub OpenQas(Control As IRibbonControl)
Shell ("C:\Program Files (x86)\QAS\QuickAddress Pro\QAPrown.exe")
' Shell "wscript W:\!Common\Templates\QAS.vbs"

End Sub
' Utility subs so we can launch the forms without the ribbon args
Public Sub LaunchTemplatePicker()
    Load frmTemplatePicker
    frmTemplatePicker.Show
End Sub

Public Sub LaunchProfileSetup(Optional SelectedUser As String = "")
    Load UserForm1
    UserForm1.Show
    
End Sub

Public Sub PCCPowerpointTemplate(Control As IRibbonControl)
    LoadPPT "PCC Presentation Template.PPTM"
End Sub

Private Sub LoadPPT(strPPT As String)
    Dim objPPT, objPresentation
    Set objPPT = CreateObject("PowerPoint.Application")
    objPPT.Visible = True
    Set objPresentation = objPPT.Presentations.Open("W:\!Common\Templates\Office_2010\Office_2010_Templates\PCC Presentation Template.PPTM")
'    Set objPresentation = objPPT.Presentations.Open(strWorkgroupTemplatesPath & strPPT)
    End Sub
