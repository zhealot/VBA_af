Attribute VB_Name = "moduleDMEventCode"
' FILE NAME: footer.dot
' FILE LOCATION: per-use Word Startup directory
' MODULE NAME: moduleDMEventCode
' Purpose:
' This module defines a custom subroutine that is called by
' DM when the user selects Insert>DM Footer from the Word menu.
' The subroutine inserts customized information into the footer.
' To make the following code work, place the following 4 lines
' of text (minus the comment (') marks) into a .reg file and
' run it.
'Windows Registry Editor Version 5.00
'
'[HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\DM_COM_Addin.WordAddin]
'"EventBeforeFooter"="moduleDMEventCode.subEventBeforeFooter"

Public Sub subEventBeforeFooter(Control As IRibbonControl)
' This subroutine is called by DM when the Insert>DM Footer
' menu item is clicked in Word.
' The DM code that normally inserts the footer is not called,
' i.e. this code replaces that code.
'
' Collect the information intended for the footer.
'Dim strFooterText As String
'strFooterText = get_footer_information()
'' Insert the information into the footer.
'insert_footer strFooterText
    If Not Trim(ActiveDocument.BuiltInDocumentProperties(wdPropertyComments).Value) = "" Then
        Call PCC_Footer
    Else
        MsgBox "No daisy footer is available yet." & vbNewLine & "Please close the document and open again in EDIT mode."
    End If
End Sub

Private Sub insert_footer(strFooterText As String)
Dim footer_selection As Selection
Dim initial_active_window_view_type As Integer
initial_active_window_view_type = ActiveWindow.View.Type
Application.EnableCancelKey = wdCancelDisabled
If (ActiveWindow.View.SplitSpecial = wdPaneNone) Then
ActiveWindow.ActivePane.View.Type = wdPageView
Else
ActiveWindow.View.Type = wdPageView
End If

ActiveWindow.View.SeekView = wdSeekMainDocument
ActiveWindow.View.SeekView = wdSeekPrimaryFooter
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Shapes.SelectAll
Set footer_selection = ActiveWindow.Selection
' Set the footer text to the sole parameter of this function.
footer_selection.Text = strFooterText
Selection.Font.Size = 8
ActiveWindow.View.Type = wdPageView
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range = footer_selection.Range
footer_selection.Paragraphs.Alignment = wdAlignParagraphLeft
ActiveWindow.View.Type = wdNormalView
ActiveWindow.View.Type = initial_active_window_view_type
Word.StatusBar = ""
Application.EnableCancelKey = wdCancelInterrupt
End Sub

Private Function get_footer_information() As String
Dim strFooterText As String
Dim objDOCSObjectsDM As Object
Dim strActiveDocumentFullName As String
Dim out_strValue As String
Dim strDocumentNumber As String
Dim strVersionNumber As String
Dim strAuthorID As String
strActiveDocumentFullName = ActiveDocument.FullName
Set objDOCSObjectsDM = CreateObject("DOCSObjects.DM")
objDOCSObjectsDM.GetDocInfo strActiveDocumentFullName, "DOCNAME", out_strValue
strDocumentName = out_strValue
objDOCSObjectsDM.GetDocInfo strActiveDocumentFullName, "AUTHOR_ID", out_strValue
strAuthorID = out_strValue
objDOCSObjectsDM.GetDocInfo strActiveDocumentFullName, "DOCNUM", out_strValue
strDocumentNumber = out_strValue
strVersionNumber = get_version_number(strActiveDocumentFullName)
strFooterText = ""
' Comment out any of the following lines to remove
' information from the footer.
' You may also re-order the lines below.
' BEGIN LINES WHICH MAY BE COMMENTED OUT OR RE-ORDERED
strFooterText = strFooterText & "PCC - #" & strDocumentNumber & "-v" & strVersionNumber & vbNewLine
' END LINES WHICH MAY BE COMMENTED OUT OR RE-ORDERED
get_footer_information = strFooterText
End Function

Private Function get_version_number(strActiveDocumentFullName As String) As String
Dim strVersionNumber As String
Dim objRegExp As Object
Dim objMatches, objMatch As Object
Dim strMatchValue As String
strVersionNumber = ""
Set objRegExp = CreateObject("VBScript.RegExp")
objRegExp.Global = True
' The following regular expression looks for the
' following pattern:
' 1 non-alphanumeric character
' followed by
' 1 letter 'v', either upper or lower case
' followed by
' 1 or more decimal digits
' followed by
' 1 non-alphanumeric character
' Examples of matching patterns:
' _v1_
' _V11_
' -V23-
' -v17-
' Thus, it catches the version number regardless of the
' delimiter used to delimit
' the various parts of a DM file name.
objRegExp.Pattern = "[^a-zA-Z0-9][vV][0-9]+[^a-zA-Z0-9]"
Set objMatches = objRegExp.Execute(strActiveDocumentFullName)
' The matches count will normally be one. For it to be
' different would generally require that the DM Administrator
' had made a change to how DM names files.
If objMatches.count = 0 Then
strVersionNumber = "<UNDEFINED>"
Else
' By default, and under normal circumstances, their
' should be only one match object. The exceptions will
' be in those situations where the author's id, the
' library name, or the document name (or some other
' value added to the document naming scheme by the
' DM Administrator) contains a string that matches
' the regular expression above (e.g. where the
' document name is "automobile_v8_engines"). If this
' is an issue, perform further expression testing to ensure
' that this function returns the version number.
Set objMatch = objMatches.item(0)
strMatchValue = objMatch.Value
strVersionNumber = Left(strMatchValue, Len(strMatchValue) - 1)
strVersionNumber = Right(strVersionNumber, Len(strVersionNumber) - 2)
End If
get_version_number = strVersionNumber
End Function


