Attribute VB_Name = "Allfields"
Const Style1 = "Instructional Text"
Const Style2 = "Instructional Text Bullets"
Public Const DOCUMENTPROPERTY = "Category" 'document property name to hold business group
Public BusinessGroup As String  'Business Group name
Public sBG As String 'short business group name
Public docA As Document

Public Sub ShowDIP(control As IRibbonControl)
    On Error Resume Next
    If Application.DisplayDocumentInformationPanel = False Then
        Application.DisplayDocumentInformationPanel = True
    Else
        Application.DisplayDocumentInformationPanel = False
    End If
End Sub

Public Sub RemoveInstructions(control As IRibbonControl)
'Search for paragraphs styled "Instructions" and delete them
    Dim rng As Range
    Dim boolFound As Boolean
    Dim boolSmartCutPaste As Boolean
    Dim MsgText As String
    
    'Smart cut & paste' setting must be false so that deleting the last paragraph
    'in a table cell doesn't change the style
    boolSmartCutPaste = Options.PasteSmartCutPaste
    Options.PasteSmartCutPaste = False
    
    On Error GoTo Bye
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(Style1)
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Execute
        If rng.End = ActiveDocument.Content.End Then GoTo NextStyle   'not whole document!
        While .Found
            boolFound = True
            rng.Delete
            If rng.Information(wdWithInTable) Then
                'if at end of cell, no paragraph to delete
                If rng.Start = rng.Cells(1).Range.End - 1 Then
                    On Error Resume Next
                    rng.Style = "TinyFont"
                    On Error GoTo Bye
                End If
            End If
            rng.SetRange Start:=rng.Start, End:=ActiveDocument.Range.End
            .Execute
        Wend
    End With
NextStyle:
    'repeat for different style
    Set rng = ActiveDocument.Content
    With rng.Find
        .Style = ActiveDocument.Styles(Style2)
        .Execute
        If rng.End = ActiveDocument.Content.End Then GoTo Bye   'not whole document!
        While .Found
            boolFound = True
            rng.Delete
            If rng.Information(wdWithInTable) Then
                'if at end of cell, no paragraph to delete
                If rng.Start = rng.Cells(1).Range.End - 1 Then
                    On Error Resume Next
                    rng.Style = "TinyFont"
                    On Error GoTo Bye
                End If
            End If
            rng.SetRange Start:=rng.Start, End:=ActiveDocument.Range.End
            .Execute
        Wend
    End With
    
Bye:
    Err.Clear
    'restore 'Smart cut & paste' setting
    Options.PasteSmartCutPaste = boolSmartCutPaste
    
    If boolFound Then
        MsgText = "All instructions have been removed."
    Else
        MsgText = "No instructions were found."
    End If
    MsgBox MsgText, vbInformation, MED
End Sub

Public Sub Logo(control As IRibbonControl)
    fmLogo.Show
End Sub

Public Function OBClick(ob As control)
    sBG = ""
    Dim ctl As control
    For Each ctl In fmLogo.frmImage.Controls
        If Right(ctl.Name, Len(ctl.Name) - 3) = Right(ob.Name, Len(ob.Name) - 2) Then
            sBG = Right(ob.Name, Len(ob.Name) - 2)
            ctl.Visible = True
        Else
            ctl.Visible = False
        End If
    Next ctl
End Function
