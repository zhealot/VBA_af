VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub cbProvider_Change()
    'keep selection and initial option buttons on all frames
    sProvider = cbProvider.Value
    If sProvider = "" Then Exit Sub
    wsGraphs.Cells(4, "B").Value = sProvider
    'reset frames
    FrameControlsEnable frmLevel, True
    FrameControlsEnable frmFieldType, True
    FrameControlsEnable frmBroad, False
    FrameControlsEnable frmNarrow, False
        
    'pupulate range
    Set rgProviderErn = FindRange(wsErn, wsErn.Columns(clmProviderErn), clmProviderErn, sProvider)
    Set rgProviderDst = FindRange(wsDst, wsDst.Columns(clmProviderDst), clmProviderDst, sProvider)
        
    'generate option buttons for frame Qualification level
    'Set frmLevel = wsGraphs.OLEObjects("frmLevel").Object
    frmLevel.Controls.Clear
    If frmLevel.Controls.Count = 0 Then
        RwLast = wsData.Cells(wsData.Rows.Count, clmLevel).End(xlUp).Row
        Set rg = wsData.Range(wsData.Cells(1, clmLevel), wsData.Cells(RwLast, clmLevel))
        For Each cl In rg
            Set ob = frmLevel.Controls.Add("Forms.OptionButton.1")
            ob.Top = (frmLevel.Controls.Count) * PosTop
            ob.Left = PosLeft
            ob.Height = OBHeight
            ob.Width = OBWidth
            ob.Caption = cl.Value
            ob.Font.Size = OBFontSize
            ob.Font.Bold = OBFontBold
            Set obLevel(cl.Row - 1).OBHandler = ob
            obLevel(cl.Row - 1).sFrame = "frmLevel"
            obLevel(cl.Row - 1).obCaption = ob.Caption
            'by defautl disable item at first
            ob.Enabled = False
        Next cl
    End If
    'validte items on frame levle
    Dim rgTmp As Range
    For Each ctr In frmLevel.Controls
        On Error Resume Next
        Set rgTmp = rgProviderDst.Find(ctr.Caption, , xlValues, xlWhole)
        If Not rgTmp Is Nothing Then
            ctr.Enabled = True
        End If
    Next ctr
    FillTable
End Sub

Private Sub Worksheet_Activate()
    'Call Init
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If obLevel(0).OBHandler Is Nothing Then
        Call Init
    End If
End Sub
