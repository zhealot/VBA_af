Attribute VB_Name = "TemplatePicker"
Function GetFileList(FileSpec As String) As Variant
'   Returns an array of filenames that match FileSpec
'   If no matching files are found, it returns False

    Dim FileArray() As Variant
    Dim FileCount As Integer
    Dim Filename As String

    On Error GoTo NoFilesFound

    FileCount = 0
    Filename = Dir(FileSpec)
    If Filename = "" Then GoTo NoFilesFound

'   Loop until no more matching files are found
    Do While Filename <> ""
        FileCount = FileCount + 1
        ReDim Preserve FileArray(1 To FileCount)
        FileArray(FileCount) = Filename
        Filename = Dir()
    Loop
    GetFileList = FileArray
    Exit Function

'   Error handler
NoFilesFound:
    GetFileList = False
End Function

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

'Function to get filetitle from full path
Public Function GetFileTitle(strFilename As String) As String
    Dim dotPos As Integer
    Dim slashPos As Integer
    Dim slashLen As Integer
    Dim dotLen As Integer
    dotPos = -1
    slashPos = 1
    For i = Len(strFilename) To 2 Step -1
      c = Mid(strFilename, i, 1)
      If -1 = dotPos And c = "." Then
        dotPos = i + 1
      ElseIf c = "\" Then
        slashPos = i + 1
        Exit For
      End If
    Next
    
    slashLen = Len(strFilename) + 1 - slashPos
    dotLen = Len(strFilename) + 2 - dotPos
    
    GetFileTitle = Mid(strFilename, slashPos, slashLen - dotLen)
End Function

Public Function GetFileExtension(strFilename As String) As String
    pos = 1
    For i = Len(strFilename) To 2 Step -1
      c = Mid(strFilename, i, 1)
      If c = "." Then
        pos = i + 1
        Exit For
      End If
    Next
    
    GetFileExtension = Mid(strFilename, pos, (Len(strFilename) + 1 - pos))
End Function

