VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Rule_Minister"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   OleObjectBlob   =   "Rule_MinisterA5_2000_Beta_UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'modified by tao@allfields.co.nz, 21/Jan/2016
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGo_Click()
ReplaceMark "Part", Me.txtPart
ReplaceMark "Amendment2", Me.txtAmendment
ReplaceMark "Amendment3", Me.txtAmendment
ReplaceMark "Part2", Me.txtPart
ReplaceMark "Part3", Me.txtPart
ReplaceMark "Part4", Me.txtPart
ReplaceMark "Amendment", Me.txtAmendment
ReplaceMark "RuleTitle", Me.txtRuleTitle
ReplaceMark "Docket", Me.txtDocket
ReplaceMark "Minister1", UCase(Me.txtMinister)
ReplaceMark "Minister2", Me.txtMinister
ReplaceMark "Title1", Me.txtMinisterTitle
ReplaceMark "Title2", Me.txtMinisterTitle

On Error Resume Next
ReplaceMark "EffectiveDate", Format(DateAdd("m", txtOffset.Text, txtDate.Text), "dd mmmm yyyy")
On Error GoTo 0


NextField
Unload Me
End Sub
Public Sub ReplaceMark(Mark As String, Value As Variant)
    ActiveDocument.Bookmarks(Mark).Range.Text = Value & ""
End Sub

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtDate.Text) Then
        MsgBox "Not a valid date"
        Cancel = True
    End If
End Sub

Private Sub txtOffset_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(txtOffset.Text) Then
        MsgBox "Not a valid number"
        Cancel = True
    End If
End Sub


Sub NextField()
'
' NextField Macro
' Macro recorded 8/11/01 by Janet Whittaker
'
    Selection.GoTo What:=wdGoToField, Which:=wdGoToNext, Count:=1, Name:=""
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
End Sub

