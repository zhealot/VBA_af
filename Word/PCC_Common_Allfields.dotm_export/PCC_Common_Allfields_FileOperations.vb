VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function PopulateCtl(IniFileName As String, cboToPopulate As Control, GroupName As String) As Boolean
On Error GoTo Error_Handler:
Dim item$
Dim count
Dim FileNo
FileNo = FreeFile
PopulateCtl = True

Open IniFileName For Input As FileNo
Do Until item$ = GroupName Or EOF(FileNo) = True
    Input #FileNo, item$
Loop
If EOF(FileNo) = True Then
    MsgBox ("The group name is incorrect!")
    PopulateCtl = False
    Exit Function
End If

'Get String from file and populate cbo
count = 0
item$ = " "
While EOF(FileNo) = False And Left(item$, FileNo) <> "[" And item$ <> ""
    Input #FileNo, item$
    If count = 0 Then
        cboToPopulate.Value = item$
    End If
    
    cboToPopulate.AddItem item$, count
    count = count + 1
    
Wend

PopulateCtl = True
Close #FileNo
Exit Function
Error_Handler:
PopulateCtl = False
'msgbox

End Function

