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
Public obNarrow(20) As New OBHandler    'container for option butoons in frmNarrow, ###capacity is manually set

'variables to keep selection of each frame/combo box
Public sProvider As String
Public sFieldType As String
Public sLevel As String
Public sBroad As String
Public sNarrow As String
'range on sheet Destinations
Public rgProviderDst As Range
Public rgLevelDst As Range
Public rgFieldTypeDst As Range
Public rgBroadDst As Range
Public rgNarrowDst As Range
'range on sheet Earnings
Public rgProviderErn As Range
Public rgLevelErn As Range
Public rgFieldTypeErn As Range
Public rgBroadErn As Range
Public rgNarrowErn As Range

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

'column in sheet Destinations
Public Const clmProviderDst = 5
Public Const clmEdumisDst = 6
Public Const clmLevelDst = 7
Public Const clmFieldTypeDst = 8
Public Const clmCodeDst = 9
Public Const clmBroadDst = 10
Public Const clmNarrowDst = 11
Public Const clmEmpDst = "L"
Public Const clmFurDst = "U"
Public Const clmBenDst = "AD"
Public Const clmOSDst = "AM"
Public Const clmOUDst = "AV"
Public Const clmNumDst = "BE"
'columns in sheet Earnings
Public Const clmProviderErn = 5
Public Const clmEdumisErn = 6
Public Const clmLevelErn = 7
Public Const clmFieldTypeErn = 8
Public Const clmCodeErn = 9
Public Const clmBroadErn = 10
Public Const clmNarrowErn = 11
Public Const clmAEErn = "M"
Public Const clmNumErn = "V"

'sheets
Public wsData As Worksheet
Public wsGraphs As Worksheet
Public wsDst As Worksheet
Public wsErn As Worksheet
'frames
Public frmLevel As MSForms.Frame
Public frmFieldType As MSForms.Frame
Public frmBroad As MSForms.Frame
Public frmNarrow As MSForms.Frame


Function Init()
'initial frames and combobox
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsGraphs = ThisWorkbook.Sheets("EOTE graphs")
    Set wsDst = ThisWorkbook.Sheets("Destinations")
    Set wsErn = ThisWorkbook.Sheets("Earnings")
    
    Dim rg As Range
    Dim RwLast As Long
    Dim cb As MSForms.ComboBox
    
    'populate Provider combo box
    Set cb = wsGraphs.OLEObjects("cbProvider").Object
    cb.Clear
    RwLast = wsData.Cells(wsData.Rows.Count, clmProvider).End(xlUp).Row
    Set rg = wsData.Range(wsData.Cells(1, clmProvider), wsData.Cells(RwLast, clmProvider))
    For Each cl In rg
        cb.AddItem cl.Value
    Next cl
    cb.Value = ""

    Set frmLevel = wsGraphs.OLEObjects("frmLevel").Object
    Set frmFieldType = wsGraphs.OLEObjects("frmFieldType").Object
    Set frmBroad = wsGraphs.OLEObjects("frmBroad").Object
    Set frmNarrow = wsGraphs.OLEObjects("frmNarrow").Object

    FrameControlsEnable frmLevel, False
    FrameControlsEnable frmFieldType, False
    FrameControlsEnable frmBroad, False
    FrameControlsEnable frmNarrow, False
        
End Function

Function FrameControlsEnable(frm As MSForms.Frame, enable As Boolean)
'enable/disable frame and controns in it
    Dim ctrl As Control
    frm.Enabled = enable
    For Each ctrl In frm.Controls
        ctrl.Enabled = enable
        ctrl.Visible = enable
        ctrl.Value = False
    Next
End Function

Public Function FindRange(ws As Worksheet, inRg As Range, clm As Long, txt As String) As Range
'search worksheet's range's column for text, return the within inRg that contains the txt in the column
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
        Loop
        Set FindRange = ws.Range(ws.Cells(RwStart, clm), ws.Cells(RwEnd - 1, ws.UsedRange.Columns.Count))
    End If
End Function

Function FillTable()
'populate the data table
    If sLevel = "" Or sProvider = "" Then Exit Function
    Select Case sFieldType
    Case "All"
        
    Case "Broad"
    
    Case "Narrow"
    
    End Select
    Application.Calculate
End Function


