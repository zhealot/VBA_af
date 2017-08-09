VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnReset_Click()
'reset all selections and forms
    Init True
    Call ClearTable
    Application.Calculate
End Sub

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


