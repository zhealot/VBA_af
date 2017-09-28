VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Address Picker"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655
   OleObjectBlob   =   "Airways Letterhead_UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    If ob1.Value Then
        sAddress = Split(AddAKL, "%")
    ElseIf ob2.Value Then
        sAddress = Split(AddWLT, "%")
    ElseIf ob3.Value Then
        sAddress = Split(AddCHC, "%")
    ElseIf ob4.Value Then
        Dim tb As MSForms.TextBox
        For i = 1 To 5
            Set tb = frmCustom.Controls("tb" & i)
            If Not CheckTB(tb) Then
                Exit Sub
            Else
                ReDim Preserve sAddress(5) As String
                sAddress(i - 1) = tb.Value
            End If
        Next i
    Else
        MsgBox "Please select an address."
        Exit Sub
    End If
    Dim bm As Bookmark
    Dim rg As Range
    For i = 1 To 5
        If ActiveDocument.Bookmarks.Exists("line" & i) Then
            Set bm = ActiveDocument.Bookmarks("line" & i)
            Set rg = bm.Range
            rg.SetRange rg.Start, rg.End - 1
            rg.Text = sAddress(i - 1)
            rg.SetRange rg.End, bm.Range.End
            rg.Text = ""
        End If
    Next i
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub ob1_Click()
    Action ob1
End Sub

Public Function Action(ob As MSForms.OptionButton)
    'text box status
    Dim blEnable As Boolean
    Dim lColor As String
    If ob.Name = "ob4" Then
        blEnable = True
        lColor = BackEnable
        'read from footer via bookmarks
        For i = 1 To 5
            If ActiveDocument.Bookmarks.Exists("line" & i) Then
                frmCustom.Controls("tb" & i).Value = ActiveDocument.Bookmarks("line" & i).Range.Text
            End If
        Next
    Else
        blEnable = False
        lColor = BackDisable
        Select Case ob.Name
        Case "ob1"
            sAddress = Split(AddAKL, "%")
        Case "ob2"
            sAddress = Split(AddWLT, "%")
        Case "ob3"
            sAddress = Split(AddCHC, "%")
        End Select
        If UBound(sAddress) = 4 Then
            For i = 0 To UBound(sAddress)
                frmCustom.Controls("tb" & i + 1).Value = sAddress(i)
            Next i
        End If
    End If
    Dim tb As MSForms.TextBox
    For i = 1 To 5
        Set tb = frmCustom.Controls("tb" & i)
        tb.Enabled = blEnable
        tb.BackColor = lColor
    Next i
End Function

Private Sub ob2_Click()
    Action ob2
End Sub

Private Sub ob3_Click()
    Action ob3
End Sub

Private Sub ob4_Click()
    Action ob4
End Sub

