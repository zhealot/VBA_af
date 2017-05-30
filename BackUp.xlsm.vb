VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim str As String
    Dim doc As Excel.Workbook
    Dim fileName As String
    Dim fn As String
    Dim tm As Long
    tm = Timer
    
    fn = Dir(ThisWorkbook.Path & "\Excel\*.xl?m", vbNormal)
    While fn <> ""
        Debug.Print fn
        If LCase(Right(fn, 4)) = "xlsm" Or LCase(Right(fn, 4)) = "xltm" Then
            Set doc = Workbooks.Open(ThisWorkbook.Path & "\Excel\" & fn, ReadOnly:=True)
            str = ThisWorkbook.Path & "\Excel\" & fn & "_export"
            MkDir str
            fileName = str & "\" & Left(doc.Name, InStrRev(doc.Name, ".") - 1)
            For i = 1 To doc.VBProject.VBComponents.Count
                doc.VBProject.VBComponents(i).Export fileName & "_" & doc.VBProject.VBComponents(i).Name & ".vb"
            Next i
            doc.Close False
            Set doc = Nothing
        End If
        fn = Dir
    Wend
    Debug.Print Timer - tm
    ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).UsedRange.Rows(ThisWorkbook.Sheets(1).UsedRange.Rows.Count).Row + 1, 1).Value = Date & ": " & Timer - tm
End Sub
