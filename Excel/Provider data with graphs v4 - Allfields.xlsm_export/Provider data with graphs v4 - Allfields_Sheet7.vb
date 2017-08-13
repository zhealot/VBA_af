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

Private Sub CommandButton1_Click()
    frmMain.Hide
    ThisWorkbook.Sheets(1).Activate
End Sub

Private Sub CommandButton2_Click()
    frmMain.Hide
    ThisWorkbook.Sheets(2).Activate
End Sub

Private Sub CommandButton3_Click()
     ActiveWindow.ScrollRow = 57
     ActiveWindow.ScrollColumn = 1
End Sub

Private Sub CommandButton4_Click()
     ActiveWindow.ScrollRow = 1
     ActiveWindow.ScrollColumn = 1
End Sub

Private Sub showForm_Click()
    frmMain.Show
End Sub


