Attribute VB_Name = "Allfields"
Public Const MAX_INI_WRITE_ATTEMPTS = 3
Public Const FILE_VERSION = "1.1"
Public Const LAST_UPDATED = "3/05/2016 3:30pm"
Public Const LAST_AUTHOR = "Allfields"
Public strWorkgroupTemplatesPath As String
Public strHelpPath As String
Public Const InstructionFile = "instructions.dotx"
Public TOOLKIT_LOADED As Boolean
Public Const ext = "dotx"
Public Const imgEx = "jpg"
Public strTemplatesPath As String
Public imgPath As String

Public Sub setDefaultTab(control As IRibbonUI)
    control.ActivateTab "TK_MinistryTemplates"
End Sub

Sub Autoexec()
    On Error Resume Next
    Dim sPath As String
    
    sPath = "C:\templates"    'setup default templates path
    If Len(ThisDocument.Paragraphs(1).Range.Text) > 4 Then
        If Right(Left(ThisDocument.Paragraphs(1).Range.Text, 3), 2) = ":\" Then
            sPath = Trim(Application.CleanString(Replace(ThisDocument.Paragraphs(1).Range.Text, Chr(13), "")))
            On Error Resume Next
            Err.Clear
            If Dir(sPath, vbDirectory) <> "" Then
                If Err.Number <> 0 Then
                    sPath = "C:\templates"
                End If
            End If
        End If
    End If
    strWorkgroupTemplatesPath = sPath
   
    sPath = "c:\Templates\Image Library" 'setup default image path
    If ThisDocument.Paragraphs.Count > 1 Then
        If Len(ThisDocument.Paragraphs(2).Range.Text) > 4 Then
            If Right(Left(ThisDocument.Paragraphs(2).Range.Text, 3), 2) = ":\" Then
                sPath = Trim(Application.CleanString(Replace(ThisDocument.Paragraphs(2).Range.Text, Chr(13), "")))
                On Error Resume Next
                Err.Clear
                If Dir(sPath, vbDirectory) <> "" Then
                    If Err.Number <> 0 Then
                        sPath = "c:\Templates\Image Library"
                    End If
                End If
            End If
        End If
    End If
    Options.DefaultFilePath(wdPicturesPath) = sPath
    imgPath = sPath
    strTemplatesPath = strWorkgroupTemplatesPath & "\"
    
    strHelpPath = strWorkgroupTemplatesPath    'set Help document path
    TOOLKIT_LOADED = True
End Sub


Public Sub ShowTemplatesMenu(control As IRibbonControl)
    Allfields.Autoexec
    CheckRequirements
    LaunchTemplatePicker
End Sub

Public Sub Images(control As IRibbonControl)
    If Not Dir(imgPath, vbDirectory) = "" Then
        Shell "explorer.exe" & " " & imgPath, vbNormalFocus
    Else
        MsgBox "Image folder not found"
    End If
End Sub

Public Sub OpenHowToGuide(control As IRibbonControl)
    If Not Dir(strHelpPath & "\" & InstructionFile) = "" Then
        Documents.Add Template:=strHelpPath & "\" & InstructionFile
    Else
        MsgBox "Help file not found"
    End If
End Sub

Public Sub ShowVersionInformation(control As IRibbonControl)
    Call MsgBox("Toolkit version " & FILE_VERSION & vbCr & vbCr _
            & "Last updated " & LAST_UPDATED & vbCr _
            & "by " & LAST_AUTHOR, vbInformation, "Template Toolkit Version Information")
End Sub

' Utility subs so we can launch the forms without the ribbon args
Public Sub LaunchTemplatePicker()
    Load frmTemplatePicker
    frmTemplatePicker.Show
End Sub

'Public Sub PCCPowerpointTemplate(control As IRibbonControl)
'    LoadPPT "PCC Presentation Template.PPTM"
'End Sub

'Private Sub LoadPPT(strPPT As String)
'    Dim objPPT, objPresentation
'    Set objPPT = CreateObject("PowerPoint.Application")
'    objPPT.Visible = True
'    'Set objPresentation = objPPT.Presentations.Open("W:\!Common\Templates\Office_2010\Office_2010_Templates\PCC Presentation Template.PPTM")
'    'Set objPresentation = objPPT.Presentations.Open(strWorkgroupTemplatesPath & "\" & strPPT)
'End Sub


Function CheckRequirements(Optional RequireUserIni As Boolean = True) As Boolean
    If Not FolderExists(strTemplatesPath) Then _
        ThrowFatalError "There does not appear to be a folder for templates." & vbCr & _
                            "Looking in: " & strTemplatesPath
End Function

Function FolderExists(fname) As Boolean
'    Dim fso
'    Set fso = CreateObject("Scripting.FileSystemObject")
    If Dir(fname, vbDirectory) = "" Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function

Sub ThrowFatalError(strError As String)
    ShowError strError
    End
End Sub

Sub ShowError(strError As String)
    MsgBox strError & vbCr & vbCr & _
            "If the problem persists contact IT Support", _
                vbCritical + vbOKOnly
End Sub

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

'Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
'
'  Dim pivot   As Variant
'  Dim tmpSwap As Variant
'  Dim tmpLow  As Long
'  Dim tmpHi   As Long
'
'  tmpLow = inLow
'  tmpHi = inHi
'
'  pivot = vArray((inLow + inHi) \ 2)
'
'  While (tmpLow <= tmpHi)
'
'     While (vArray(tmpLow) < pivot And tmpLow < inHi)
'        tmpLow = tmpLow + 1
'     Wend
'
'     While (pivot < vArray(tmpHi) And tmpHi > inLow)
'        tmpHi = tmpHi - 1
'     Wend
'
'     If (tmpLow <= tmpHi) Then
'        tmpSwap = vArray(tmpLow)
'        vArray(tmpLow) = vArray(tmpHi)
'        vArray(tmpHi) = tmpSwap
'        tmpLow = tmpLow + 1
'        tmpHi = tmpHi - 1
'     End If
'
'  Wend
'
'  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
'  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
'
'End Sub

'Function to get filetitle from full path
Public Function GetFileTitle(strFilename As String) As String
    Dim dotPos As Integer
    Dim slashPos As Integer
    Dim slashLen As Integer
    Dim dotLen As Integer
    dotPos = -1
    slashPos = 1
    For i = Len(strFilename) To 2 Step -1
      C = Mid(strFilename, i, 1)
      If -1 = dotPos And C = "." Then
        dotPos = i + 1
      ElseIf C = "\" Then
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
      C = Mid(strFilename, i, 1)
      If C = "." Then
        pos = i + 1
        Exit For
      End If
    Next
    GetFileExtension = Mid(strFilename, pos, (Len(strFilename) + 1 - pos))
End Function

Function ListSubFolders(fld As String) As Object
'return 1 level subfolers in folder
    Dim FileSystem As Object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set ListSubFolders = FileSystem.GetFolder(fld)
End Function

Function HasFileType(fld As String, ext As String) As Boolean
'check if folder has file(s) of extension, include subfolers
    HasFileType = False
    If Not FolderExists(fld) Then Exit Function
    If Dir(fld & "\*." & ext) <> "" Then
        HasFileType = True
        Exit Function
    End If
    Dim f
    Dim objFld As Object
    Set objFld = ListSubFolders(fld)
    If objFld Is Nothing Then Exit Function
    For Each f In objFld.SubFolders
        If HasFileType(f.Path, ext) Then
            HasFileType = True
            Exit Function
        End If
    Next f
End Function

Function ColourValue(cl As String) As Long
    Select Case LCase(Trim(cl))
        Case "exaqua"
            ColourValue = 10333797
        Case "exblue"
            ColourValue = 12953687
        Case "exgreen"
            ColourValue = 3915194
        Case "exred"
            ColourValue = 2382313
        Case "exrose"
            ColourValue = 7759822
        Case "inblue"
            ColourValue = 14458112
        Case "ingreen"
            ColourValue = 3850658
        Case "inpink"
            ColourValue = 7422434
        Case "inyellow"
            ColourValue = 39423
        Case Else
            ColourValue = 0
        End Select
End Function

