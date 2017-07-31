VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OBhandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const clmNarrow = 7     'column holding narrow fields
Public WithEvents OBHandler As MSForms.OptionButton
Attribute OBHandler.VB_VarHelpID = -1
Public obCaption As String  'option button caption
Public sFrame As String     'indicate which frame the option button comes from

Private Sub OBHandler_Click()
    Dim wsGraphs As Worksheet
    Set wsGraphs = ThisWorkbook.Sheets("EOTE graphs")
    Dim frmLevel As MSForms.Frame
    Dim frmBroad As MSForms.Frame
    Dim frmNarrow As MSForms.Frame
    Dim frmFieldType As MSForms.Frame
    Dim iRw As Long
    
    Set frmLevel = wsGraphs.OLEObjects("frmLevel").Object
    Set frmNarrow = wsGraphs.OLEObjects("frmNarrow").Object
    Set frmBroad = wsGraphs.OLEObjects("frmBroad").Object
    Set frmNarrow = wsGraphs.OLEObjects("frmNarrow").Object
    
    Select Case sFrame
    'click in Qualification Level
    Case "frmLevel"
        sLevel = OBHandler.Caption
        wsGraphs.Cells(5, "B").Value = sLevel
        'setup relevant ranges
        Set rgLevelDst = FindRange(wsDst, rgProviderDst, clmLevelDst, sLevel)
        Set rgLevelErn = FindRange(wsErn, rgProviderErn, clmLevelErn, sLevel)
        'generate option buttons for frame field type level
        Set frmFieldType = wsGraphs.OLEObjects("frmFieldType").Object
        frmFieldType.Controls.Clear
        If frmFieldType.Controls.Count = 0 Then
            RwLast = wsData.Cells(wsData.Rows.Count, clmFieldType).End(xlUp).Row
            Set rg = wsData.Range(wsData.Cells(1, clmFieldType), wsData.Cells(RwLast, clmFieldType))
            For Each cl In rg
                Set ob = frmFieldType.Controls.Add("Forms.OptionButton.1")
                ob.Top = (frmFieldType.Controls.Count) * PosTop
                ob.Left = PosLeft
                ob.Height = OBHeight
                ob.Width = OBWidth
                ob.Caption = cl.Value
                ob.Font.Size = OBFontSize
                ob.Font.Bold = OBFontBold
                Set obFieldType(cl.Row - 1).OBHandler = ob
                obFieldType(cl.Row - 1).sFrame = "frmFieldType"
                obFieldType(cl.Row - 1).obCaption = ob.Caption
            Next cl
        End If

    'click in Field type
    Case "frmFieldType"
        sFieldType = OBHandler.Caption
        wsGraphs.Cells(6, "B").Value = sFieldType
        sNarrow = ""
        sBroad = ""
        'setup broad and narrow frames
        Select Case OBHandler.Caption
            Case "All"
                FrameControlsEnable frmBroad, False
                FrameControlsEnable frmNarrow, False
                'setup relevant ranges
                Set rgFieldTypeDst = FindRange(wsDst, rgLevelDst, clmFieldTypeDst, sFieldType)
                Set rgFieldTypeErn = FindRange(wsErn, rgLevelErn, clmFieldTypeErn, sFieldType)
                FillTable
            Case "Broad", "Narrow"
                If OBHandler.Caption = "Broad" Then
                    FrameControlsEnable frmBroad, True
                    FrameControlsEnable frmNarrow, False
                Else
                    FrameControlsEnable frmBroad, True
                    FrameControlsEnable frmNarrow, True
                End If
                frmNarrow.Controls.Clear
                'use same range as level
                Set rgFieldTypeDst = rgLevelDst
                Set rgFieldTypeErn = rgLevelErn
                'generate controls for frame Broad field of study
                Set frmBroad = wsGraphs.OLEObjects("frmBroad").Object
                frmBroad.Controls.Clear
                If frmBroad.Controls.Count = 0 Then
                    RwLast = wsData.Cells(wsData.Rows.Count, clmBroad).End(xlUp).Row
                    Set rg = wsData.Range(wsData.Cells(1, clmBroad), wsData.Cells(RwLast, clmBroad))
                    For Each cl In rg
                        Set ob = frmBroad.Controls.Add("Forms.OptionButton.1")
                        ob.Top = (frmBroad.Controls.Count) * PosTop
                        ob.Left = PosLeft
                        ob.Height = OBHeight
                        ob.Width = OBWidth
                        ob.Caption = cl.Value
                        ob.Font.Size = OBFontSize
                        ob.Font.Bold = OBFontBold
                        Set obBroad(cl.Row - 1).OBHandler = ob
                        obBroad(cl.Row - 1).obCaption = ob.Caption
                        obBroad(cl.Row - 1).sFrame = "frmBroad"
                        'disable item first
                        ob.Enabled = False
                        ob.Value = False
                    Next cl
                End If
                'validte items in frame Broad
                For iRw = 0 To rgFieldTypeDst.Rows.Count - 1
                    If wsDst.Cells(rgFieldTypeDst.Row + iRw, clmFieldTypeDst).Value = "Broad" Then
                        For Each ctr In frmBroad.Controls
                            On Error Resume Next
                            If ctr.Caption = wsDst.Cells(rgFieldTypeDst.Row + iRw, clmBroadDst).Value Then
                                ctr.Enabled = True
                                Exit For
                            End If
                        Next ctr
                    End If
                Next
            Case Else
        End Select
        
    'click in Broad Field study
    Case "frmBroad"
        sBroad = OBHandler.Caption
        wsGraphs.Cells(8, "B").Value = sBroad
        sNarrow = ""
        'setup relevant ranges
        'within a certain level range, the broad field name leads to broad field rows only
        'so just search broad field name in the level range will get the correct rows
        'below rgFieldTypeDst = rgLevelDst, rgFieldTypeErn = rgFieldTypeErn
        Set rgBroadDst = FindRange(wsDst, rgFieldTypeDst, clmBroadDst, sBroad)
        Set rgBroadErn = FindRange(wsErn, rgFieldTypeErn, clmBroadErn, sBroad)
        'populate Narrow fields
        If sFieldType = "Narrow" Then
            FrameControlsEnable frmNarrow, True
            frmNarrow.Controls.Clear
            If frmNarrow.Controls.Count = 0 Then
                Set wsData = ThisWorkbook.Sheets("Data")
                'set range based on button clicked in frmBroad
                Set rg = wsData.Columns(clmBroadAll).Find(Replace(Replace(OBHandler.Caption, " ", ""), ",", ""), LookIn:=xlValues)
                If Not rg Is Nothing Then
                    'define rg to next row and next column
                    Set rg = rg.Offset(1, 1)
                    Set cl = rg
                    'search for next empty cell
                    Do While Len(cl.Value) <> 0
                        Set cl = cl.Offset(1, 0)
                    Loop
                    'redefine rg
                    Set rg = wsData.Range(rg, cl.Offset(-1, 0))
                End If
                Set cl = Nothing
                'populate option buttons for narrow field study
                For i = 1 To rg.Cells.Count
                    Set ob = frmNarrow.Controls.Add("Forms.OptionButton.1")
                    ob.Top = frmNarrow.Controls.Count * PosTop
                    ob.Left = PosLeft
                    ob.Height = OBHeight
                    ob.Width = OBWidth + 50
                    ob.Caption = rg.Cells(i).Value
                    ob.Font.Size = OBFontSize
                    ob.Font.Bold = OBFontBold
                    Set obNarrow(i).OBHandler = ob
                    obNarrow(i).sFrame = "frmNarrow"
                    obNarrow(i).obCaption = ob.Caption
                    'set option to disable at first
                    ob.Enabled = False
                    ob.Value = False
                Next i
                frmNarrow.Height = rg.Cells.Count * OBHeight + 30
                'validate items on frame Narrow
                For iRw = 1 To rgBroadDst.Rows.Count - 1
                    For Each ctr In frmNarrow.Controls
                        On Error Resume Next
                        If ctr.Caption = wsDst.Cells(rgBroadDst.Row + iRw, clmNarrowDst).Value Then
                            ctr.Enabled = True
                            Exit For
                        End If
                    Next ctr
                Next
            End If
        Else
            FrameControlsEnable frmNarrow, False
            FillTable
        End If
        
    'click in Narrow field study
    Case "frmNarrow"
        sNarrow = OBHandler.Caption
        wsGraphs.Cells(9, "B").Value = sNarrow
        Set rgNarrowDst = FindRange(wsDst, rgBroadDst, clmNarrowDst, sNarrow)
        Set rgNarrowErn = FindRange(wsErn, rgBroadErn, clmNarrowErn, sNarrow)
        FillTable
    Case Else
    End Select
End Sub
