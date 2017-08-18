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
    Calculate
    If Not TriggerEvent Then Exit Sub
'    sProvider = wsGraphs.Cells(56, 10).Value
'    sLevel = GetValue(frmLevel)
'    sFieldType = GetValue(frmFieldType)
'    sBroad = GetValue(frmBroad)
'    sNarrow = GetValue(frmNarrow)
    Select Case sFrame
    'click in Qualification Level
    Case "frmLevel"
        Call LevelClick(OBHandler)
    'click in Field type
    Case "frmFieldType"
        Call FieldTypeClick(OBHandler)
    'click in Broad Field study
    Case "frmBroad"
        Call BroadClick(OBHandler)
    'click in Narrow field study
    Case "frmNarrow"
        Call NarrowClick(OBHandler)
    Case Else
    End Select
    Calculate
End Sub

