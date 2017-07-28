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
Public obCaption As String
Public sFrame As String     'indicate which frame the option button comes from

Private Sub OBHandler_Click()
    Dim wsGraphs As Worksheet
    Set wsGraphs = ThisWorkbook.Sheets("EOTE graphs")
    Dim obFrmNarrow As MSForms.Frame
    Set obFrmNarrow = wsGraphs.OLEObjects("frmNarrow").Object
    'buttons in frmBroad clicked
    If sFrame = "frmBroad" Then
        obFrmNarrow.Controls.Clear
        If obFrmNarrow.Controls.Count = 0 Then
            Dim wsData As Worksheet
            Set wsData = ThisWorkbook.Sheets("Data")
            '###figure out set range based on button clicked in frmBroad
            
            
            RwLast = wsData.Cells(wsData.Rows.Count, clmNarrow).End(xlUp).Row
            Dim rg As Range
            Set rg = wsData.Range(wsData.Cells(1, clmNarrow), wsData.Cells(RwLast, clmNarrow))
            Dim cl As Range
            For Each cl In rg
                Set ob = obFrmNarrow.Controls.Add("Forms.OptionButton.1")
                ob.Top = obFrmNarrow.Controls.Count * PosTop
                ob.Left = PosLeft
                ob.Height = OBHeight
                ob.Width = OBWidth
                ob.Caption = cl.Value
                ob.Font.Size = OBFontSize
                ob.Font.Bold = OBFontBold
                Set obFieldType(cl.Row - 1).OBHandler = ob
                obFieldType(cl.Row - 1).sFrame = "frmNarrow"
                obFieldType(cl.Row - 1).obCaption = ob.Caption
            Next cl
        End If
    End If
End Sub
