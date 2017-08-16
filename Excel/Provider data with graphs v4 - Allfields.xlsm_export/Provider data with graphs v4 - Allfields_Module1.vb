Attribute VB_Name = "Module1"
'-----------------------------------------------------------------------------
' Developed for TEC
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             August 2017
' Description:      optimized user interface of making selection, faster
'                   respond time.
'-----------------------------------------------------------------------------

Public obLevel(7) As New OBHandler      'container for option buttons in frmLevel
Public obFieldType(2) As New OBHandler  'container for option butoons in frmFieldType
Public obBroad(11) As New OBHandler     'container for option buttons in frmBroad
Public obNarrow(30) As New OBHandler    'container for option butoons in frmNarrow, ###capacity is manually set
'event trigger flag
Public TriggerEvent As Boolean          'flasg to trigger option button event
Public TriggerProvider As Boolean       'flag to trigger combo box Provider
Public TriggerSimilar As Boolean        'flag to trigger combo box Similar

'variables to keep selection of each frame/combo box
Public sProvider As String
Public sFieldType As String
Public sLevel As String
Public sBroad As String
Public sNarrow As String

'const for frame layout
Public Const PosTop = 20              'control top position
Public Const PosLeft = 5             'control left position
Public Const OBHeight = 20             'option button height
Public Const OBWidth = 250              'option butoon width
Public Const OBFontBold = False         'font bold
Public Const OBFontSize = 10           'font size
'constants for column in sheet data
Public Const clmLevel = 1              'column of qulification levels
Public Const clmFieldType = 2          'column of field type
Public Const clmBroad = 4              'column of broad field of study
Public Const clmBroadNum = 3           'column of broad field code
Public Const clmBroadAll = 6
Public Const clmNarrow = 7             'column of narrow field of study
Public Const clmNarrowNum = 8          '###column of narrow field code
Public Const clmProvider = 9            'column of provider names

'data table position
Public Const ProviderBaseRw = 58        'first cell of provider data
Public Const AllBaseRw = 68             'first cell of all provider data
Public Const BaseClm = 2
Public Const ProviderRw = 57             'cell that keeps provider name in table
Public Const ProviderClm = "A"

'column in sheet Destinations
Public Const clmProviderDst = "E"
Public Const clmLevelDst = "G"
Public Const clmFieldTypeDst = "H"
Public Const clmBroadDst = "J"
Public Const clmNarrowDst = "K"
Public Const clmEmpDst = "L"
Public Const clmFurDst = "U"
Public Const clmBenDst = "AD"
Public Const clmOSDst = "AM"
Public Const clmOUDst = "AV"
Public Const clmNumDst = "BE"
'range on sheet Destinations
Public rgProviderDst As Range
Public rgLevelDst As Range
Public rgFieldTypeDst As Range
Public rgBroadDst As Range
Public rgNarrowDst As Range
Public rgProviderDstAll As Range    'range for 'All provider' data
Public rgLevelDstAll As Range       '
Public rgFieldTypeDstAll As Range   '
Public rgBroadDstAll As Range       '
Public rgNarrowDstAll As Range      '

'columns in sheet Earnings
Public Const clmProviderErn = "E"
Public Const clmLevelErn = "G"
Public Const clmFieldTypeErn = "H"
Public Const clmBroadErn = "J"
Public Const clmNarrowErn = "K"
Public Const clmPercentile = "L"
Public Const clmAEErn = "M"
Public Const clmNumErn = "V"
'range on sheet Earnings
Public rgProviderErn As Range
Public rgLevelErn As Range
Public rgFieldTypeErn As Range
Public rgBroadErn As Range
Public rgNarrowErn As Range
Public rgProviderErnAll As Range    'range for 'All provider' data
Public rgLevelErnAll As Range       '
Public rgFieldTypeErnAll As Range   '
Public rgBroadErnAll As Range       '
Public rgNarrowErnAll As Range      '
Public rgNonAllDst As Range         'range of non-all provider on sheet Destinations

'sheets
Public wsData As Worksheet
Public wsGraphs As Worksheet
Public wsDst As Worksheet
Public wsErn As Worksheet
'frames and combo boxes
Public frmLevel As MSForms.Frame
Public frmFieldType As MSForms.Frame
Public frmBroad As MSForms.Frame
Public frmNarrow As MSForms.Frame
Public cbSimilar As MSForms.ComboBox

Function Init(all As Boolean)
'initial frames and combobox
    On Error Resume Next
    TriggerEvent = True
    TriggerProvider = True
    TriggerSimilar = True
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsGraphs = ThisWorkbook.Sheets("Graphs and Tables")
    Set wsDst = ThisWorkbook.Sheets("Destinations")
    Set wsErn = ThisWorkbook.Sheets("Earnings")
    'setup 'All provider' range
    Set rgProviderDstAll = wsDst.Range(wsDst.Cells(3, clmProviderDst), wsDst.Cells(587, wsDst.UsedRange.Columns.Count))
    Set rgProviderErnAll = wsErn.Range(wsErn.Cells(3, clmProviderErn), wsErn.Cells(1730, wsErn.UsedRange.Columns.Count))
    'hook up frames
    Set frmLevel = frmMain.frmLevel     'wsGraphs.OLEObjects("frmLevel").Object
    Set frmFieldType = frmMain.frmFieldType     'wsGraphs.OLEObjects("frmFieldType").Object
    Set frmBroad = frmMain.frmBroad     'wsGraphs.OLEObjects("frmBroad").Object
    Set frmNarrow = frmMain.frmNarrow       'wsGraphs.OLEObjects("frmNarrow").Object
    Set cbSimilar = frmMain.cbSimilar   'wsGraphs.OLEObjects("cbSimilar").Object
    
    wsGraphs.Activate
    Call ClearTable
    frmMain.Show
    Dim rg As Range
    Dim RwLast As Long
    Dim cb As MSForms.ComboBox
    Dim ob As MSForms.OptionButton
    
    CancelFilter
    
    'populate frame Levels
    frmLevel.Controls.Clear
    RwLast = wsData.Cells(wsData.Rows.Count, clmLevel).End(xlUp).Row
    Set rg = wsData.Range(wsData.Cells(1, clmLevel), wsData.Cells(RwLast, clmLevel))
    For Each cl In rg
        Set ob = frmLevel.Controls.Add("Forms.OptionButton.1")
        ob.Top = (frmLevel.Controls.Count) * PosTop
        ob.Left = PosLeft
        ob.Height = OBHeight
        ob.Width = OBWidth
        ob.Caption = cl.Value
        ob.Name = cl.Value
        ob.Font.Size = OBFontSize
        ob.Font.Bold = OBFontBold
        Set obLevel(cl.Row - 1).OBHandler = ob
        obLevel(cl.Row - 1).obCaption = ob.Caption
        obLevel(cl.Row - 1).sFrame = "frmLevel"
        'disable item first
        ob.Enabled = False
        ob.Value = False
    Next cl
    frmLevel.Height = rg.Cells.Count * OBHeight + 30
    
    'populate frame FieldType
    frmFieldType.Controls.Clear
    RwLast = wsData.Cells(wsData.Rows.Count, clmFieldType).End(xlUp).Row
    Set rg = wsData.Range(wsData.Cells(1, clmFieldType), wsData.Cells(RwLast, clmFieldType))
    For Each cl In rg
        Set ob = frmFieldType.Controls.Add("Forms.OptionButton.1")
        ob.Top = (frmFieldType.Controls.Count) * PosTop
        ob.Left = PosLeft
        ob.Height = OBHeight
        ob.Width = OBWidth
        ob.Caption = cl.Value & " field of study" '###15/08/2017
        ob.Name = cl.Value
        ob.Font.Size = OBFontSize
        ob.Font.Bold = OBFontBold
        Set obFieldType(cl.Row - 1).OBHandler = ob
        obFieldType(cl.Row - 1).obCaption = ob.Name 'ob.Caption ###15/08/2017
        obFieldType(cl.Row - 1).sFrame = "frmFieldType"
        'disable item first
        ob.Enabled = False
        ob.Value = False
    Next cl
    frmFieldType.Height = rg.Cells.Count * OBHeight + 30
    
    'binding complete, check whether need to initialize all
    If all Then
        'reset all variables and frames
        sLevel = ""
        sFieldType = ""
        sBroad = ""
        sNarrow = ""
        'populate Provider combo box
        Set cb = frmMain.cbProvider       'wsGraphs.OLEObjects("cbProvider").Object
        If cb.ListCount = 0 Then
            cb.Clear
            RwLast = wsData.Cells(wsData.Rows.Count, clmProvider).End(xlUp).Row
            Set rg = wsData.Range(wsData.Cells(1, clmProvider), wsData.Cells(RwLast, clmProvider))
            For Each cl In rg
                cb.AddItem cl.Value
            Next cl
        End If
        cb.Value = ""
        Call cbProvicderChange
        Set cb = frmMain.cbSimilar       'wsGraphs.OLEObjects("cbSimilar").Object
        cb.Clear
        cb.Value = ""
        FrameControlsEnable frmLevel, False
        FrameControlsEnable frmFieldType, False
        FrameControlsEnable frmBroad, False
        FrameControlsEnable frmNarrow, False
    Else
        Exit Function
    End If
    Application.Calculate
End Function

Function FrameControlsEnable(frm As MSForms.Frame, enable As Boolean)
'enable/disable frame and controns in it
    On Error Resume Next
    Dim ctrl As Control
    frm.Enabled = enable
    TriggerEvent = False
    For Each ctrl In frm.Controls
        ctrl.Enabled = enable
        'ctrl.Visible = enable
        ctrl.Value = False
    Next
    TriggerEvent = True
End Function

Public Function FindRange(ws As Worksheet, inRg As Range, clm As String, txt As String) As Range
'search worksheet's range's column for text, return the within inRg that contains the txt in the column
    On Error Resume Next
    If inRg Is Nothing Or txt = "" Then Exit Function
    'pupulate range
    Dim rgTmp As Range  'temporary range
    Dim RwStart As Long 'start row number
    Dim RwEnd As Long   'end row number
    Set rgTmp = Nothing
    'find first row of the select provider
    Set rgTmp = inRg.Find(txt, , xlValues, xlWhole)
    If rgTmp Is Nothing Then
        Set FindRange = Nothing
        Exit Function
    Else
        RwStart = rgTmp.Row
        RwEnd = rgTmp.Row
        Do While ws.Cells(RwEnd, clm).Value = txt
            RwEnd = RwEnd + 1
            'check current row still within inRg
            If RwEnd > inRg.Row + inRg.Rows.Count - 1 Then
                Exit Do
            End If
        Loop
        Set FindRange = ws.Range(ws.Cells(RwStart, inRg.Column), ws.Cells(RwEnd - 1, ws.UsedRange.Columns.Count))
    End If
End Function

Function FillTable()
'populate the data table
    On Error Resume Next
    Application.ScreenUpdating = False
    If sLevel = "" Or sProvider = "" Then Exit Function
    Dim baseClDst As Range
    Dim baseClDstAll As Range
    Dim baseClErn As Range
    Dim baseClErnAll As Range
    Dim clSrc As Range  'source cell
    Dim clTbl As Range  'table cell
    
    Call GetFrameValues
    
    If sFieldType = "" Then
        Call ClearTable
        
        Exit Function
    End If

    Select Case sFieldType
    Case "All"
        Set baseClDst = FindRange(wsDst, rgLevelDst, clmFieldTypeDst, "All")
        Set baseClDstAll = FindRange(wsDst, rgLevelDstAll, clmFieldTypeDst, "All")
        Set baseClErn = FindRange(wsErn, rgLevelErn, clmFieldTypeErn, "All")
        Set baseClErn = FindRange(wsErn, baseClErn, clmPercentile, "Median")            'search for the 'Median' row
        Set baseClErnAll = FindRange(wsErn, rgLevelErnAll, clmFieldTypeErn, "All")
        Set baseClErnAll = FindRange(wsErn, baseClErnAll, clmPercentile, "Median")      'search for the 'Median' row
    Case "Broad"
        If sBroad = "" Then
            Call ClearTable
            Exit Function
        End If
        Set baseClDst = FindRange(wsDst, rgBroadDst, clmFieldTypeDst, "Broad")
        Set baseClDstAll = FindRange(wsDst, rgBroadDstAll, clmFieldTypeDst, "Broad")
        Set baseClErn = FindRange(wsErn, rgBroadErn, clmFieldTypeErn, "Broad")
        Set baseClErn = FindRange(wsErn, rgBroadErn, clmPercentile, "Median")
        Set baseClErnAll = FindRange(wsErn, rgBroadErnAll, clmFieldTypeErn, "Broad")
        Set baseClErnAll = FindRange(wsErn, rgBroadErnAll, clmPercentile, "Median")
    Case "Narrow"
        If sNarrow = "" Then
            Call ClearTable
            Exit Function
        End If
        Set baseClDst = FindRange(wsDst, rgNarrowDst, clmNarrowDst, sNarrow)
        Set baseClDstAll = FindRange(wsDst, rgNarrowDstAll, clmNarrowDst, sNarrow)
        Set baseClErn = FindRange(wsErn, rgNarrowErn, clmNarrowErn, sNarrow)
        Set baseClErn = FindRange(wsErn, rgNarrowErn, clmPercentile, "Median")
        Set baseClErnAll = FindRange(wsErn, rgNarrowErnAll, clmNarrowErn, sNarrow)
        Set baseClErnAll = FindRange(wsErn, rgNarrowErnAll, clmPercentile, "Median")
    End Select
    'get one row only for Destinations data, then fill the Provider table
    If baseClDst.Rows.Count = 1 Then
        'fill Provider table 'Destinations' part
        Set clSrc = wsDst.Cells(baseClDst.Row, clmNumDst)
        Set clTbl = wsGraphs.Cells(ProviderBaseRw, BaseClm)
        ExtractCell wsDst, clSrc, wsGraphs, clTbl, 9, 9         'fill 'Number of graduates' row
        Set clSrc = wsDst.Cells(baseClDst.Row, clmEmpDst)
        Set clTbl = wsGraphs.Cells(ProviderBaseRw + 1, BaseClm) 'fill other rows
        ExtractCell wsDst, clSrc, wsGraphs, clTbl, 45, 9
        'fill Provider table 'Earnings' part
        Set clSrc = wsErn.Cells(baseClErn.Row, clmAEErn)
        Set clTbl = wsGraphs.Cells(ProviderBaseRw + 7, BaseClm)
        ExtractCell wsErn, clSrc, wsGraphs, clTbl, 9, 9
        Set clSrc = wsErn.Cells(baseClErn.Row, clmNumErn)
        Set clTbl = wsGraphs.Cells(ProviderBaseRw + 6, BaseClm)
        ExtractCell wsErn, clSrc, wsGraphs, clTbl, 9, 9
        
        'fill All Provider table 'Destinations' part
        Set clSrc = wsDst.Cells(baseClDstAll.Row, clmNumDst)
        Set clTbl = wsGraphs.Cells(AllBaseRw, BaseClm)
        ExtractCell wsDst, clSrc, wsGraphs, clTbl, 9, 9         'fill 'Number of graduates' row
        Set clSrc = wsDst.Cells(baseClDstAll.Row, clmEmpDst)
        Set clTbl = wsGraphs.Cells(AllBaseRw + 1, BaseClm) 'fill other rows
        ExtractCell wsDst, clSrc, wsGraphs, clTbl, 45, 9
        'fill All Provider table 'Earnings' part
        Set clSrc = wsErn.Cells(baseClErnAll.Row, clmAEErn)
        Set clTbl = wsGraphs.Cells(AllBaseRw + 7, BaseClm)
        ExtractCell wsErn, clSrc, wsGraphs, clTbl, 9, 9
        Set clSrc = wsErn.Cells(baseClErnAll.Row, clmNumErn)
        Set clTbl = wsGraphs.Cells(AllBaseRw + 6, BaseClm)
        ExtractCell wsErn, clSrc, wsGraphs, clTbl, 9, 9
    End If
    Application.Calculate
    ThisWorkbook.RefreshAll
    Application.ScreenUpdating = True
    Calculate
End Function

Function ExtractCell(wsSrc As Worksheet, clSrc As Range, wsDst As Worksheet, clDst As Range, SrcCellNum As Integer, DstCellNum As Integer)
'copy cells value from a row in source to destinate cells
'   wsSrc: source worksheet
'   clSrc: base cell in source worksheet
'   wsDst: destinate worksheet
'   clDst: base cell in destinate worksheet
'   SrcCellNum: how many cells to be copied from source
'   DstCellNum: how many cells in each row in destinate.
'   eg. SrcCellNum = 27, DstCellNum =12, then there will be copied into 3 rows with last row has 3 cells.
    On Error Resume Next
    If wsSrc Is Nothing Or _
        clSrc Is Nothing Or _
        wsDst Is Nothing Or _
        clDst Is Nothing Then Exit Function
        
    If SrcCellNum < 1 Or DstCellNum < 1 Then Exit Function
    
    Dim iSrc As Long
'    '=================
'    Dim SrcRwCur As Long
'    Dim SrcClmCur As Long
'    Dim DstRwCur As Long
'    Dim DstClmCur As Long
'    '=================
    
    For iSrc = 0 To SrcCellNum - 1
'        DstRwCur = clDst.Row + (iSrc \ DstCellNum)
'        DstClmCur = clDst.Column + (iSrc Mod DstCellNum)
'        SrcRwCur = clSrc.Row
'        SrcClmCur = clSrc.Column + iSrc
        wsDst.Cells(clDst.Row + (iSrc \ DstCellNum), clDst.Column + (iSrc Mod DstCellNum)).Value = wsSrc.Cells(clSrc.Row, clSrc.Column + iSrc).Value
    Next iSrc
End Function

Function ClearTable()
'clear provider and all provider data table
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim rg As Range
    If wsGraphs.Cells(ProviderBaseRw, BaseClm).Value <> "" Then
        Set rg = wsGraphs.Range(wsGraphs.Cells(ProviderBaseRw, BaseClm), wsGraphs.Cells(ProviderBaseRw + 7, BaseClm + 8))
        rg.Value = ""
    End If
    If wsGraphs.Cells(AllBaseRw, BaseClm).Value <> "" Then
        Set rg = wsGraphs.Range(wsGraphs.Cells(AllBaseRw, BaseClm), wsGraphs.Cells(AllBaseRw + 7, BaseClm + 8))
        rg.Value = ""
    End If
    Application.ScreenUpdating = True
End Function

Function GetFrameValues()
'get each frame selected values if it's enabled
    On Error Resume Next
    If frmLevel Is Nothing Then
        sLevel = ""
    Else
        If frmLevel.Enabled Then
            sLevel = ""
            For Each ctr In frmLevel.Controls
                If ctr.Value And ctr.Enabled Then
                    sLevel = ctr.Caption
                    Exit For
                End If
            Next ctr
        Else
            sLevel = ""
        End If
    End If
    
    If frmFieldType Is Nothing Then
        sFieldType = ""
    Else
        If frmFieldType.Enabled Then
            sFieldType = ""
            For Each ctr In frmFieldType.Controls
                If ctr.Value And ctr.Enabled Then
                    sFieldType = ctr.Name   'ctr.Caption, ###15/08/2017
                    Exit For
                End If
            Next ctr
        Else
            sFieldType = ""
        End If
    End If
    
    If frmBroad Is Nothing Then
        sBroad = ""
    Else
        If frmBroad.Enabled Then
            sBroad = ""
            For Each ctr In frmBroad.Controls
                If ctr.Value And ctr.Enabled Then
                    sBroad = ctr.Caption
                    Exit For
                End If
            Next ctr
        Else
            sBroad = ""
        End If
    End If
    
    If frmNarrow Is Nothing Then
        sNarrow = ""
    Else
        If frmBroad.Enabled Then
            sNarrow = ""
            For Each ctr In frmNarrow.Controls
                If ctr.Value And ctr.Enabled Then
                    sNarrow = ctr.Caption
                    Exit For
                End If
            Next ctr
        Else
            sNarrow = ""
        End If
    End If
End Function

Function LevelClick(ob As MSForms.OptionButton)
    On Error Resume Next
    sLevel = ob.Caption
    'setup relevant ranges
    CancelFilter
    Set rgLevelDst = FindRange(wsDst, rgProviderDst, clmLevelDst, sLevel)
    Set rgLevelErn = FindRange(wsErn, rgProviderErn, clmLevelErn, sLevel)
    Set rgLevelDstAll = FindRange(wsDst, rgProviderDstAll, clmLevelDst, sLevel)
    Set rgLevelErnAll = FindRange(wsErn, rgProviderErnAll, clmLevelErn, sLevel)
    
    'generate option buttons for frame field type level
    If frmFieldType.Controls.Count = 0 Then
        RwLast = wsData.Cells(wsData.Rows.Count, clmFieldType).End(xlUp).Row
        Set rg = wsData.Range(wsData.Cells(1, clmFieldType), wsData.Cells(RwLast, clmFieldType))
        For Each cl In rg
            Set ob = frmFieldType.Controls.Add("Forms.OptionButton.1")
            ob.Top = (frmFieldType.Controls.Count) * PosTop
            ob.Left = PosLeft
            ob.Height = OBHeight
            ob.Width = OBWidth
            ob.Name = cl.Value  '###15/08/2017
            ob.Caption = cl.Value & " field of study"   '###15/08/2017
            ob.Font.Size = OBFontSize
            ob.Font.Bold = OBFontBold
        Next cl
    End If
    'heritage status from previous stage
    If sFieldType <> "" Then
        Dim opTmp As MSForms.OptionButton
        Set opTmp = Nothing
        For Each ctr In frmFieldType.Controls
            'If ctr.Caption = sFieldType Then   ###15/08/2017
            If ctr.Name = sFieldType Then
                TriggerEvent = False
                Set opTmp = ctr
                opTmp.Value = True
                TriggerEvent = True
                Exit For
            End If
        Next
        '### when setting its value, it triggers the OBHandler event, so below call is not necessary
        Call FieldTypeClick(opTmp)
    Else
        'try fill table, actually will clear the table
        FillTable
    End If
    
End Function

Function FieldTypeClick(ob As MSForms.OptionButton)
    On Error Resume Next
    sFieldType = ob.Name 'ob.Caption ###15/08/2017
    CancelFilter
    'setup broad and narrow frames
    Select Case ob.Name    'ob.Caption  ###15/08/2017
        Case "All"
            FrameControlsEnable frmBroad, False
            FrameControlsEnable frmNarrow, False
            'setup relevant ranges
            Set rgFieldTypeDst = FindRange(wsDst, rgLevelDst, clmFieldTypeDst, sFieldType)
            Set rgFieldTypeErn = FindRange(wsErn, rgLevelErn, clmFieldTypeErn, sFieldType)
            Set rgFieldTypeDstAll = FindRange(wsDst, rgLevelDstAll, clmFieldTypeDst, sFieldType)
            Set rgFieldTypeErnAll = FindRange(wsErn, rgLevelErnAll, clmFieldTypeErn, sFieldType)
            TriggerSimilar = False
            frmMain.cbSimilar.Clear
            TriggerSimilar = True
            FillTable
        Case "Broad", "Narrow"
            'If ob.Caption = "Broad" Then   ###15/08/2017
            If ob.Name = "Broad" Then
                FrameControlsEnable frmBroad, True
                FrameControlsEnable frmNarrow, False
            Else
                FrameControlsEnable frmBroad, True
                FrameControlsEnable frmNarrow, True
            End If
            'If ob.Caption = "Broad" Then   ###15/08/2017
            If ob.Name = "Braod" Then
                frmNarrow.Controls.Clear
            End If
            'use same range as level
            Set rgFieldTypeDst = rgLevelDst
            Set rgFieldTypeErn = rgLevelErn
            Set rgFieldTypeDstAll = rgLevelDstAll
            Set rgFieldTypeErnAll = rgLevelErnAll
            
            'generate controls for frame Broad field of study
            Set frmBroad = frmMain.frmBroad 'wsGraphs.OLEObjects("frmBroad").Object
            'frmBroad.Controls.Clear
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
            SetOBs frmBroad, False  'disabled all first, then enable the ones in range
            For iRw = 0 To rgFieldTypeDst.Rows.Count - 1
                If wsDst.Cells(rgFieldTypeDst.Row + iRw, clmFieldTypeDst).Value = "Broad" Then
                    For i = 0 To frmBroad.Controls.Count - 1
                        If frmBroad.Controls(i).Caption = wsDst.Cells(rgFieldTypeDst.Row + iRw, clmBroadDst).Value Then
                            Dim oppp As MSForms.OptionButton
                            Set oppp = frmBroad.Controls(i)
                            oppp.Enabled = True
                            Set obBroad(i).OBHandler = oppp
                            obBroad(i).sFrame = "frmBroad"
                            Exit For
                        End If
                    Next i
                End If
            Next iRw
            'select the option button that has been selected from previous stage
            If sBroad <> "" Then
                Dim obtmp As MSForms.OptionButton
                Set obtmp = Nothing
                For Each ctr In frmBroad.Controls
                    If ctr.Caption = sBroad And ctr.Enabled Then
                        TriggerEvent = False
                        Set obtmp = ctr
                        obtmp.Value = True
                        TriggerEvent = True
                        Exit For
                    End If
                Next ctr
                'no Broad item is selected
                If obtmp Is Nothing Then
                    sNarrow = ""
                    frmNarrow.Controls.Clear
                Else
                    Call BroadClick(obtmp)
                End If
            Else
                FillTable
            End If
        Case Else
    End Select
End Function

Function BroadClick(ob As MSForms.OptionButton)
    On Error Resume Next
    sBroad = ob.Caption
    CancelFilter
    'setup relevant ranges
    'within a certain level range, the broad field name leads to broad field rows only
    'so just search broad field name in the level range will get the correct rows
    'below rgFieldTypeDst = rgLevelDst, rgFieldTypeErn = rgFieldTypeErn
    Set rgBroadDst = FindRange(wsDst, rgFieldTypeDst, clmBroadDst, sBroad)
    Set rgBroadErn = FindRange(wsErn, rgFieldTypeErn, clmBroadErn, sBroad)
    Set rgBroadDstAll = FindRange(wsDst, rgFieldTypeDstAll, clmBroadDst, sBroad)
    Set rgBroadErnAll = FindRange(wsErn, rgFieldTypeErnAll, clmBroadErn, sBroad)
            
    'populate Narrow fields
    If sFieldType = "Narrow" Then
        FrameControlsEnable frmNarrow, True
        frmNarrow.Controls.Clear
        If frmNarrow.Controls.Count = 0 Then
            Set wsData = ThisWorkbook.Sheets("Data")
            'set range based on button clicked in frmBroad
            Set rg = wsData.Columns(clmBroadAll).Find(Replace(Replace(ob.Caption, " ", ""), ",", ""), LookIn:=xlValues)
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
            If sNarrow <> "" Then
                Dim obNarr As MSForms.OptionButton
                Set obNarr = Nothing
                For Each ctr In frmNarrow.Controls
                    If ctr.Caption = sNarrow And ctr.Enabled Then
                        TriggerEvent = False
                        Set obNarr = ctr
                        obNarr.Value = True
                        TriggerEvent = True
                        Exit For
                    End If
                Next
                If Not obNarr Is Nothing Then
                    Call NarrowClick(obNarr)
                End If
            End If
        End If
    Else
        FrameControlsEnable frmNarrow, False
        FillTable
        ListSimilarProviders
    End If
End Function

Function NarrowClick(ob As MSForms.OptionButton)
    On Error Resume Next
    sNarrow = ob.Caption
    CancelFilter
    Set rgNarrowDst = FindRange(wsDst, rgBroadDst, clmNarrowDst, sNarrow)
    Set rgNarrowErn = FindRange(wsErn, rgBroadErn, clmNarrowErn, sNarrow)
    Set rgNarrowDstAll = FindRange(wsDst, rgBroadDstAll, clmNarrowDst, sNarrow)
    Set rgNarrowErnAll = FindRange(wsErn, rgBroadErnAll, clmNarrowErn, sNarrow)
    FillTable
    ListSimilarProviders
End Function

Function cbProvicderChange()
    On Error Resume Next
    CancelFilter
    'keep selection and initial option buttons on all frames
    Dim cbProvider As MSForms.ComboBox
    Set cbProvider = frmMain.cbProvider     'wsGraphs.OLEObjects("cbProvider").Object
    sProvider = cbProvider.Value
    wsGraphs.Cells(ProviderRw, ProviderClm).Value = sProvider
    Application.Calculate   'refresh
    If sProvider = "" Then Exit Function
    If wsErn Is Nothing Then
        Call Init(False)
    End If
    'reset frames
    FrameControlsEnable frmLevel, True
    FrameControlsEnable frmFieldType, True
    FrameControlsEnable frmBroad, False
    FrameControlsEnable frmNarrow, False
    
    'pupulate range
    Set rgProviderErn = FindRange(wsErn, wsErn.Columns(clmProviderErn), clmProviderErn, sProvider)
    Set rgProviderDst = FindRange(wsDst, wsDst.Columns(clmProviderDst), clmProviderDst, sProvider)
        
    'generate option buttons for frame Qualification level
    'frmLevel.Controls.Clear
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
        Next cl
    End If
    'validte items on frame levle
    Dim rgTmp As Range
    Dim iAry As Integer
    Dim ctr As Control
    For i = 0 To frmLevel.Controls.Count - 1
        On Error Resume Next
        Set ctr = frmLevel.Controls(i)
        Set rgTmp = rgProviderDst.Find(ctr.Caption, , xlValues, xlWhole)
        If rgTmp Is Nothing Then
            ctr.Enabled = False
        Else
            ctr.Enabled = True
        End If
        Set obLevel(i).OBHandler = ctr
        obLevel(i).sFrame = "frmLevel"
    Next i
    
    'check if level has valid status
    If sLevel <> "" Then
        Dim obLevelTmp As MSForms.OptionButton
        Set obLevelTmp = Nothing
        For Each ctr In frmLevel.Controls
            If ctr.Caption = sLevel Then
                If ctr.Enabled Then
                    TriggerEvent = False
                    Set obLevelTmp = ctr
                    obLevelTmp.Value = True
                    TriggerEvent = True
                Else
                    sLevel = ""
                    TriggerSimilar = False
                    cbSimilar.Value = ""
                    TriggerSimilar = True
                End If
                Exit For
            End If
        Next ctr
        If Not obLevelTmp Is Nothing Then
            Call LevelClick(obLevelTmp)
        End If
    End If
    FillTable
End Function

Function SetOBs(frm As MSForms.Frame, status As Boolean)
'set option buttons within a frame enabled or disabled
    For Each ctr In frm.Controls
        ctr.Enabled = status
    Next ctr
End Function

Private Sub showForm(ribbon As IRibbonControl)
    frmMain.Show
End Sub

Function ListSimilarProviders()
'list all providers that have the same field of study
    If sProvider = "" Or sFieldType = "All" Or sBroad = "" Then Exit Function
    
    Dim RwFirst As Long 'first found range
    Dim rgSearching As Range    'other found
    Dim sTxt As String
    Dim clm As String
    
    If sNarrow = "" Then
        sTxt = sBroad
        clm = clmBroadDst
    Else
        sTxt = sNarrow
        clm = clmNarrowDst
    End If
    'find out similar providers by Broad field
    Set rgSearching = wsDst.Columns(clm).Find(sTxt, , xlValues, xlWhole)
    TriggerSimilar = False
    cbSimilar.Clear
    TriggerSimilar = True
    'find first occurance and keep the row number
    If Not rgSearching Is Nothing Then
        RwFirst = rgSearching.Row
    Else
        Exit Function
    End If
    Do
        'check different Provider name
        If wsDst.Cells(rgSearching.Row, clmProviderDst).Value <> "All providers" Then
            'check same Level and same FieldType
            If wsDst.Cells(rgSearching.Row, clmLevelDst).Value = sLevel _
                And wsDst.Cells(rgSearching.Row, clmFieldTypeDst).Value = sFieldType _
                And wsDst.Cells(rgSearching.Row, clmBroadDst).Value = sBroad Then
                'check has values other then 'Numbers'
                Dim cl As Range
                For Each cl In wsDst.Range(wsDst.Cells(rgSearching.Row, clmEmpDst), wsDst.Cells(rgSearching.Row, "BD"))
                    If cl.Value <> "S" And cl.Value <> "" Then
                        cbSimilar.AddItem wsDst.Cells(rgSearching.Row, clmProviderDst).Value, cbSimilar.ListCount
                        Exit For
                    End If
                Next cl
            End If
        End If
        Set rgSearching = wsDst.Columns(clm).FindNext(rgSearching)
    Loop While rgSearching.Row > RwFirst
    SortCb cbSimilar
    TriggerSimilar = False
    cbSimilar.Value = sProvider
    TriggerSimilar = True
End Function

Public Function SortCb(cb As MSForms.ComboBox)
'sort combo box items
    If cb.ListCount < 2 Then Exit Function
    Dim vList As Variant
    Dim vTmp As Variant
    vList = cb.List
    Dim i As Long
    Dim j As Long
    For i = LBound(vList, 1) To UBound(vList, 1)
        For j = i + 1 To UBound(vList, 1)
            If vList(i, 0) > vList(j, 0) Then
                vTmp = vList(i, 0)
                vList(i, 0) = vList(j, 0)
                vList(j, 0) = vTmp
            End If
        Next
    Next
    cb.Clear
    For i = LBound(vList, 1) To UBound(vList, 1)
        cb.AddItem vList(i, 0)
    Next i
End Function

Public Function CancelFilter()
'show all data in sheet destinations and ernings
    If wsDst.AutoFilterMode Then
        wsDst.AutoFilterMode = False
    End If
    If wsErn.AutoFilterMode Then
        wsErn.AutoFilterMode = False
    End If
End Function
