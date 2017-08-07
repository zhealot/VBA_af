VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub cbProvider_Change()
    Call cbProvicderChange
End Sub

Private Sub cbSimilar_Change()
    If TriggerSimilar Then
        cbProvider.Value = cbSimilar.Value
    End If
End Sub

Private Sub Worksheet_Activate()
    'Call Init
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If cbProvider.ListCount = 0 Or obFieldType(0).OBHandler Is Nothing Then
        Call Init(False)
    End If
End Sub

