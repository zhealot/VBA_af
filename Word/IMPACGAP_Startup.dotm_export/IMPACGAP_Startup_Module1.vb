Attribute VB_Name = "Module1"
Option Explicit
Public DPICoefficient As Double
Dim oAppClass As New ThisApplication
'get windows DPI
Private Const LOGPIXELSX As Long = 88
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare PtrSafe Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Function GetDpi() As Long
    Dim iDPI As Long
    Dim hdcScreen As Long
    iDPI = -1
    hdcScreen = GetDC(0)
    If hdcScreen Then
        iDPI = GetDeviceCaps(hdcScreen, LOGPIXELSX)
        ReleaseDC 0, hdcScreen
    End If
    GetDpi = iDPI
End Function

Public Sub AutoExec()
    On Error Resume Next
    Set oAppClass.oApp = Word.Application
    DPICoefficient = 1
    DPICoefficient = 96 / GetDpi
End Sub
