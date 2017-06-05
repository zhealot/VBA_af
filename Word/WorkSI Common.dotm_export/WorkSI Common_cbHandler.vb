VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const COLOUR_SELECTED = &HFF0000    'font colour for selected checkbox
Const COLOUR_NONSELECTED = &H0&     'font colour for non-selected checkbox

Public WithEvents cb As MSForms.CheckBox
Attribute cb.VB_VarHelpID = -1
Public Caption As String

Private Sub cb_Click()
    'checkbox ticked before and not current one
    If cb.Value = False And cb.Font.Bold = False Then
        cb.Value = True
    End If
    Dim i As Integer
    For i = 0 To UBound(cbSections)
        Dim tmpCB As cbHandler
        Set tmpCB = cbSections(i)
        tmpCB.cb.Font.Bold = False
        tmpCB.cb.ForeColor = COLOUR_NONSELECTED
    Next i
    'set checkbox font colour
    cb.Font.Bold = cb.Value
    cb.ForeColor = IIf(cb.Value, COLOUR_SELECTED, COLOUR_NONSELECTED)
    'generate checkboxes into TEMPLATES frame
    fmMain.fmTemplates.Controls.Clear
    If cb.Value Then
        Dim blk As Block
        Dim ckbCounter As Integer
        ckbCounter = 0
        For i = 0 To UBound(Blocks)
            Set blk = Blocks(i)
            'not get '10' group mixed up with '1' group
            If Left(blk.Name, 2) = cb.Tag Or (Left(blk.Name, 1) & " " = Left(cb.Tag, 2) And Left(blk.Name, 2) <> "10") Then
                Dim ckb As MSForms.CheckBox
                Set ckb = fmMain.fmTemplates.Controls.Add("Forms.Checkbox.1", "cb1")
                With ckb
                    .Top = TOP_GAP * ckbCounter + 12
                    .Left = LEFT_COLUMN_1
                    .Width = WIDTH_SECTION * 2
                    .Height = HEIGHT_SECTION
                    .Font.Name = FONT_NAME
                    .Font.Size = FONT_SIZE
                    .Caption = blk.Name
                    .Tag = Left(.Caption, 2)
                    .Value = Blocks(i).Selected
                End With
                ReDim Preserve cbSelected(ckbCounter)
                Set cbSelected(ckbCounter).cb = ckb
                ckbCounter = ckbCounter + 1
            End If
        Next i
    End If
    'set frame scroll bar to work
    If ckbCounter * TOP_GAP > fmMain.fmTemplates.Height Then
        fmMain.fmTemplates.ScrollBars = fmScrollBarsVertical
        With fmMain.fmTemplates
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = ckbCounter * TOP_GAP + 10
        End With
    Else
        'enable scroll bar to scroll to top
        fmMain.fmTemplates.ScrollBars = fmScrollBarsVertical
        If fmMain.fmTemplates.Controls.Count > 0 Then
            fmMain.fmTemplates.ScrollTop = True
        End If
        fmMain.fmTemplates.ScrollBars = fmScrollBarsNone
    End If
End Sub
