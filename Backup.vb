Private Sub Document_Open()
    Dim str As String
    Dim doc As Document
    Dim filename As String
    Dim FN As String
    Dim tm As Long
    tm = Timer
    
    FN = Dir(ThisDocument.Path & "\word\*.do?m", vbNormal)
    While FN <> ""
        Debug.Print FN
        If LCase(Right(FN, 4)) = "dotm" Or LCase(Right(FN, 4) = "docm") Then
            On Error Resume Next
            Set doc = Documents.Open(ThisDocument.Path & "\word\" & FN, ReadOnly:=True, Visible:=False)
            str = ThisDocument.Path & "\word\" & FN & "_export"
            MkDir str
            filename = str & "\" & Left(doc.Name, InStrRev(doc.Name, ".") - 1)
            For i = 1 To doc.VBProject.VBComponents.Count
                doc.VBProject.VBComponents(i).Export filename & "_" & doc.VBProject.VBComponents(i).Name & ".vb"
            Next i
            If doc.Path & "\" & doc.Name <> ThisDocument.Path & "\" & ThisDocument.Name Then
                doc.Close False
                Set doc = Nothing
            Else
                Exit Sub
            End If
        End If
        FN = Dir
    Wend
    Debug.Print Timer - tm
End Sub


