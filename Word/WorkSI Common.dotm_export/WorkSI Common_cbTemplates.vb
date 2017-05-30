VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cbTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents cb As MSForms.CheckBox
Attribute cb.VB_VarHelpID = -1
Public Caption As String

Private Sub cb_Click()
    fmMain.fmSelected.Controls.Clear
    Dim iCounter As Integer
    For i = 0 To UBound(Blocks)
        If cb.Caption = Blocks(i).Name Then
            Blocks(i).Selected = cb.Value
            Exit For
        End If
    Next i
    For i = 0 To UBound(Blocks)
        If Blocks(i).Selected Then
            Dim tb As MSForms.Label
            Set tb = fmMain.fmSelected.Controls.Add("Forms.Label.1", "tb1")
            With tb
                .Top = TOP_GAP_LABEL * iCounter + 10
                .Left = LEFT_LABEL
                .Width = WIDTH_LABEL
                .Height = HEIGHT_LABEL
                .Font.Name = FONT_NAME
                .Font.Size = FONT_SIZE_LABEL
                .Caption = IIf(Blocks(i).Description = "", Blocks(i).Name, Blocks(i).Description)
            End With
            '####TBD: add to a array
            iCounter = iCounter + 1
        End If
    Next i
    'set frame scroll bar to work
    If iCounter * TOP_GAP_LABEL > fmMain.fmSelected.Height Then
        fmMain.fmSelected.ScrollBars = fmScrollBarsVertical
        With fmMain.fmSelected
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = iCounter * TOP_GAP_LABEL + 10
        End With
    Else
        fmMain.fmSelected.ScrollBars = fmScrollBarsNone
    End If
End Sub
