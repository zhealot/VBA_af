VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public IsIMPAC As Boolean
Public WithEvents oApp As Word.Application
Attribute oApp.VB_VarHelpID = -1

'Private Sub oApp_DocumentOpen(ByVal Doc As Document)
'    CheckIMPAC
'End Sub
'
'Private Sub oApp_NewDocument(ByVal Doc As Document)
'    CheckIMPAC
'End Sub

Private Sub oApp_WindowSelectionChange(ByVal Sel As Selection)
    fmColour.Hide
'    If Not IsIMPAC Then Exit Sub
    On Error GoTo NoMember
    If Sel.Tables.Count > 0 Then
        If Sel.Tables(1).Columns.Count > 1 Then
            If Not Left(Sel.Tables(1).Cell(1, Sel.Tables(1).Columns.Count).Range.Text, 8) = "PRIORITY" Then Exit Sub
            If Sel.Cells.Count = 1 Then
                If Sel.Cells(1).ColumnIndex = Sel.Tables(1).Columns.Count Then
                    Dim t As Long
                    Dim l As Long
                    Dim h As Long
                    Dim w As Long
                    ActiveWindow.GetPoint l, t, w, h, Sel.Range
                    'take in DPI coefficient to adjust form location
                    fmColour.Top = Application.PixelsToPoints(t) * DPICoefficient + 15
                    fmColour.Left = Application.PixelsToPoints(l) * DPICoefficient + 5
                    fmColour.Show
                End If
            End If
        End If
    End If
NoMember:
End Sub
'
'Public Function CheckIMPAC()
'    'check if a IMPAC document
'    IsIMPAC = False
'    On Error Resume Next
'    Dim ProType As Integer
'    ProType = ActiveDocument.ProtectionType
'    If ProType <> wdNoProtection Then
'        ActiveDocument.Unprotect
'    End If
'    Dim rg As Range
'    Set rg = ActiveDocument.Range
'    rg.Collapse wdCollapseStart
'    With rg.Find
'        .ClearFormatting
'        .Text = "HEALTH & SAFETY ASSESSMENT"
'        '.MatchCase = True
'        .Wrap = wdFindStop
'        .Forward = True
'        .Execute
'        If .Found Then
'            If rg.Information(wdActiveEndPageNumber) = 1 Then
'                IsIMPAC = True
'            End If
'        End If
'    End With
'    ActiveDocument.Protect ProType
'End Function
