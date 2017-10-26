Attribute VB_Name = "Ribbon"
' **************************************************
' Ribbon Control Module for PCC template toolkit.
' Responsible for delegating button-presses on the ribbon to
' the respective macros.
' @author Patrick, Allfields Customised Solutions, 04 978 7101
' ***************************************************.


Public Sub ShowTemplatesMenu(control As IRibbonControl)
    FilePaths.Autoexec
    CheckRequirements
    LaunchTemplatePicker
End Sub

Public Sub ShowProfileSetup(control As IRibbonControl)
    FilePaths.Autoexec
    If Not CheckRequirements(True) Then _
        LaunchProfileSetup
End Sub

Public Sub OpenHowToGuide(control As IRibbonControl)
    Documents.Add Template:="Instructions.dotx"
End Sub

Public Sub ShowVersionInformation(control As IRibbonControl)
    Call MsgBox("Toolkit version " & FILE_VERSION & vbCr & vbCr _
            & "Last updated " & LAST_UPDATED & vbCr _
            & "by " & LAST_AUTHOR, vbInformation, "Template Toolkit Version Information")
End Sub

Public Sub OpenQas(control As IRibbonControl)
Shell ("C:\Program Files (x86)\QAS\QuickAddress Pro\QAPrown.exe")
' Shell "wscript W:\!Common\Templates\QAS.vbs"
End Sub


Public Sub subEventBeforeFooter(control As IRibbonControl)
    If Not Trim(ActiveDocument.BuiltInDocumentProperties(wdPropertyComments).Value) = "" Then
        Call PCC_Footer
    Else
        MsgBox "No daisy number is available yet." & vbNewLine & "Please close the document and open again in EDIT mode."
    End If
End Sub

Public Sub ShowhideDaisy(control As IRibbonControl)
    Dim doc As Document
    Set doc = ActiveDocument
    Dim hd As HeaderFooter
    Dim i As Integer
    
    For i = 1 To doc.Sections.count
        If doc.Sections(i).Headers(wdHeaderFooterEvenPages).Exists Then
            Set hd = doc.Sections(i).Headers(wdHeaderFooterEvenPages)
            If InStr(hd.Range.Paragraphs.First.Range.Text, "Ref") > 0 Then
                hd.Range.Paragraphs.First.Range.Font.ColorIndex = IIf(hd.Range.Paragraphs.First.Range.Font.ColorIndex = wdWhite, wdAuto, wdWhite)
            End If
        End If
        If doc.Sections(i).Headers(wdHeaderFooterFirstPage).Exists Then
            Set hd = doc.Sections(i).Headers(wdHeaderFooterFirstPage)
            If InStr(hd.Range.Paragraphs.First.Range.Text, "Ref") > 0 Then
                hd.Range.Paragraphs.First.Range.Font.ColorIndex = IIf(hd.Range.Paragraphs.First.Range.Font.ColorIndex = wdWhite, wdAuto, wdWhite)
            End If
        End If
        If doc.Sections(i).Headers(wdHeaderFooterPrimary).Exists Then
            Set hd = doc.Sections(i).Headers(wdHeaderFooterPrimary)
            If InStr(hd.Range.Paragraphs.First.Range.Text, "Ref") > 0 Then
                hd.Range.Paragraphs.First.Range.Font.ColorIndex = IIf(hd.Range.Paragraphs.First.Range.Font.ColorIndex = wdWhite, wdAuto, wdWhite)
            End If
        End If
    Next i
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

Public Sub PCCPowerpointTemplate(control As IRibbonControl)
    LoadPPT "PowerPoint template.pptx"
End Sub

Private Sub LoadPPT(strPPT As String)
    Dim objPPT, objPresentation
    Set objPPT = CreateObject("PowerPoint.Application")
    objPPT.Visible = True
    FilePaths.Autoexec
    '### Set objPresentation = objPPT.Presentations.Open("W:\!Common\Templates\Office_2010\Office_2010_Templates\PCC Presentation Template.PPTM")
    Set objPresentation = objPPT.Presentations.Open(strWorkgroupTemplatesPath & "\" & strPPT)
    End Sub
