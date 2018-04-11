VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim sPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .InitialFileName = ThisDocument.Path & "\"
        .ButtonName = "OK"
        If .Show = -1 Then
            sPath = .SelectedItems(1) & "\"
        End If
    End With
    Call FixLogos(sPath)
End Sub

Private Sub Document_Open()
    On Error Resume Next
    Application.DisplayDocumentInformationPanel = True
    Call Allfields.AutoExec
End Sub

