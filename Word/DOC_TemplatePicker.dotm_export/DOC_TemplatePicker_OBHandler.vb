VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OBHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents cb As MSForms.OptionButton
Attribute cb.VB_VarHelpID = -1
Public Caption As String
Public Group As String
Private Sub cb_Click()
    Dim ImagePath As String
    ImagePath = imgPath & "\" & Group & IIf(Group <> "", "-", "") & Caption & "." & imgEx
    If Dir(ImagePath, vbNormal) <> "" Then
        frmTemplatePicker.imgPreview.Picture = LoadPicture(ImagePath, frmTemplatePicker.imgPreview.Width, frmTemplatePicker.imgPreview.Height)
    End If
End Sub
