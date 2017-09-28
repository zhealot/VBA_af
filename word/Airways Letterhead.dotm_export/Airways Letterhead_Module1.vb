Attribute VB_Name = "Module1"
Public Const BackEnable = &H80000005
Public Const BackDisable = &HEFEFEF
Public Const AddWLT = "Level 7 | Majestic Centre%100 Willis Street | Wellington 6011%PO Box 294 | Wellington 6140%New Zealand%Phone +64 4 471 1888"
Public Const AddAKL = "Level 2 | 6 Leonard Isitt Drive%Auckland Airport | Auckland 2022%PO Box 53093 | Auckland Airport%Auckland 2150 | New Zealand%Phone +64 4 471 1888"
Public Const AddCHC = "26 Sir William Pickering Drive%Russley | Christchurch 8053%PO Box 14131 | Christchurch 8544%New Zealand%Phone +64 4 471 1888"
Public sSelected As String
Public ADDCUS As String
Public sAddress() As String

Sub Address(control As IRibbonControl)
    UserForm1.Show
End Sub

Public Function CheckTB(tb As MSForms.TextBox) As Boolean
    CheckTB = True
    If Trim(tb.Text) = "" Then
        MsgBox "Please enter text here."
        tb.SetFocus
        CheckTB = False
    End If
End Function
