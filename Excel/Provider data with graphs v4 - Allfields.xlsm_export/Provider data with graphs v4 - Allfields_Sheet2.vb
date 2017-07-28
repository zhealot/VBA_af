VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const PosTop = 20              'control top position
Const PosLeft = 5             'control left position
Const OBHeight = 20             'option button height
Const OBWidth = 250              'option butoon width
Const OBFontBold = False         'font bold
Const OBFontSize = 10           'font size
Const clmLevel = 1              'column of qulification levels
Const clmFieldType = 2          'column of field type
Const clmBroad = 4              'column of broad field of study
Const clmBroadNum = 3           'column of broad field code
Const clmNarrow = 7             'column of narrow field of study
Const clmNarrowNum = 8          '###column of narrow field code


Private Sub Worksheet_Activate()
'try to generate all buttons for frames
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Dim rg As Range
    Dim cl As Range
    Dim RwLast As Long
    Dim ob As MSForms.OptionButton
    'generate option buttons for frame Qualification level
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
        Next cl
    End If
    
    'generate option buttons for frame field type level
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
    
    'generate controls for frame Broad field of study
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
        Next cl
    End If
End Sub

Sub testss()
    Dim ob As Object
    Set ob = ActiveSheet.OLEObjects("frmBroad")
    Debug.Print ob.Name
    Debug.Print ob.Object.Controls.Count
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'Dim Cell As Range
'    For Each Cell In Target
'        If Cell.Address = "$B$6" Then
'            Application.EnableEvents = False
'                Range("B8:B9").ClearContents
'            Application.EnableEvents = True
'            End If
'    Next Cell
'    For Each Cell In Target
'        If Cell.Address = "$B$8" Then
'            Application.EnableEvents = False
'                Range("B9").ClearContents
'            Application.EnableEvents = True
'            End If
'    Next Cell
'    For Each Cell In Target
'        If Cell.Address = "$B$40" Then
'            Application.EnableEvents = False
'                Range("B42:B43").ClearContents
'            Application.EnableEvents = True
'            End If
'    Next Cell
'    For Each Cell In Target
'        If Cell.Address = "$B$42" Then
'            Application.EnableEvents = False
'                Range("B43").ClearContents
'            Application.EnableEvents = True
'            End If
'    Next Cell
End Sub

Sub test()
    
End Sub
