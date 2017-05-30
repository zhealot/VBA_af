VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOffset 
   Caption         =   "Offset Setting"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3195
   OleObjectBlob   =   "Rule_MinisterA5_2000_Beta_frmOffset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOK_Click()
    If Not IsNumeric(tbNumber.Text) Then
        MsgBox "Please enter a number"
        tbNumber.SetFocus
        Exit Sub
    Else
        If Not Int(tbNumber.Text) = tbNumber.Text Then
            MsgBox "Please enter an integer"
            tbNumber.SetFocus
            Exit Sub
        End If
    End If
    
    Me.hide
    
    Dim cc As ContentControl
    Dim doc As Document
    Dim strTmp As String
    Dim strInput As String
    Dim strOffsetUnit As String
    
    Set doc = ActiveDocument
    If obDay.Value Then
        strOffsetUnit = "d"
    Else
        If obMonth.Value Then
            strOffsetUnit = "m"
        Else
            If obYear.Value Then
                strOffsetUnit = "y"
            Else
                MsgBox "Please choose an offset unit"
                Exit Sub
            End If
        End If
    End If
    
    On Error Resume Next
    Set cc = doc.ContentControls.Add(wdContentControlText, Selection.Range)
    If Err.Number > 0 Then
        MsgBox "Not able to insert date here"
        Exit Sub
    End If
    On Error GoTo 0
    cc.Title = "OffsetDate"
    cc.Tag = strOffsetUnit & tbNumber.Text
    On Error Resume Next
    strTmp = doc.SelectContentControlsByTitle("EffectiveDate").Item(1).Range.Text
    If IsDate(strTmp) Then
        If strOffsetUnit = "y" Then
            strOffsetUnit = "yyyy"
        End If
        strTmp = Format(DateAdd(strOffsetUnit, tbNumber.Text, strTmp), DateFormat)
    End If
    cc.Range.Text = strTmp
End Sub

Private Sub CommandButton2_Click()
    Me.hide
End Sub

Private Sub CommandButton3_Click()
    Me.hide
End Sub

