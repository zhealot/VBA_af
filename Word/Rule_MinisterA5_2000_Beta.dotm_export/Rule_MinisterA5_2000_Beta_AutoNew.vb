Attribute VB_Name = "AutoNew"
Public Const DateFormat = "d MMMM yyyy"

Sub AutoNew()
       UserForm1.Show
End Sub

Sub insertOffsetDate(control As IRibbonControl)
    If ActiveDocument.SelectContentControlsByTitle("EffectiveDate").Count > 0 Then
        If Not IsDate(ActiveDocument.SelectContentControlsByTitle("EffectiveDate").Item(1).Range.Text) Then
            MsgBox "Effective date is not correctly set." & vbNewLine & "Not able to insert offset date now."
        Else
            frmOffset.Show
        End If
    Else
        MsgBox "Content Control 'EffectiveDate' cannot be found"
    End If
End Sub
