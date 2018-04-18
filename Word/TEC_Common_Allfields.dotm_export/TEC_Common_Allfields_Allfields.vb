Attribute VB_Name = "Allfields"
Const TEMPLATE_FOLDER = "\Templates"
Public Const DEFAULT_FOLDER = "TAS Templates"
Public strWorkgroupTemplatesPath As String
Public strHelpPath As String
Public Const InstructionFile = "Help Document.docx" '### instruction docu name
Public TOOLKIT_LOADED As Boolean
Public Const ext = "dotx,pptx,potx,xltx"
Public cTypes As New Collection
Public sStr As Variant
Public Const sDelimiter = ","
Public Const imgEx = "jpg"
Public strTemplatesPath As String
Public imgPath As String

Public Sub setDefaultTab(control As IRibbonUI)
    control.ActivateTab "TK_MinistryTemplates"
End Sub

Sub Autoexec()
    On Error Resume Next
    If TOOLKIT_LOADED = False Then
        Dim spath As String
        Dim i As Integer
        Dim extA() As String
        extA = Split(ext, sDelimiter)
        On Error Resume Next
        For i = LBound(extA) To UBound(extA)
            cTypes.Add extA(i), extA(i)
        Next
        
        strWorkgroupTemplatesPath = Application.Options.DefaultFilePath(wdWorkgroupTemplatesPath) & TEMPLATE_FOLDER
        imgPath = Application.Options.DefaultFilePath(wdWorkgroupTemplatesPath) & "\Images"  '### picture folder location
        strHelpPath = Application.Options.DefaultFilePath(wdWorkgroupTemplatesPath) & "\Help File"    'set Help document path
        strTemplatesPath = strWorkgroupTemplatesPath & "\"
        TOOLKIT_LOADED = True
    End If
End Sub


Public Sub ShowTemplatesMenu(control As IRibbonControl)
    Allfields.Autoexec
    CheckRequirements
    LaunchTemplatePicker
End Sub

Public Sub Images(control As IRibbonControl)
    Allfields.Autoexec
    If Not Dir(imgPath, vbDirectory) = "" Then
        Shell "explorer.exe" & " " & imgPath, vbNormalFocus
    Else
        MsgBox "Image folder not found"
    End If
End Sub

Public Sub OpenHowToGuide(control As IRibbonControl)
    Allfields.Autoexec
    If Not Dir(strHelpPath & "\" & InstructionFile) = "" Then
        Dim doc As Word.Document
        Set doc = Word.Documents.Add(Template:=strHelpPath & "\" & InstructionFile)
        doc.Activate
    Else
        MsgBox "Help file not found"
    End If
End Sub

' Utility subs so we can launch the forms without the ribbon args
Public Sub LaunchTemplatePicker()
    Load frmTemplatePicker
    frmTemplatePicker.Show
End Sub


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


Function ListSubFolders(fld As String) As Object
'return 1 level subfolers in folder
    Dim FileSystem As Object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set ListSubFolders = FileSystem.GetFolder(fld)
End Function

Function HasFileType(fld As String, types As Collection) As Boolean
'check if folder has file(s) of extension, include subfolers
    HasFileType = False
    'check folder root
    If Not FolderExists(fld) Then Exit Function
    Dim s As Variant
    For Each s In types
        If Dir(fld & "\*." & s) <> "" Then
            HasFileType = True
            Exit Function
        End If
    Next s
    
    'check subfolder
    Dim f
    Dim objFld As Object
    Set objFld = ListSubFolders(fld)
    If objFld Is Nothing Then Exit Function
    For Each f In objFld.subfolders
        For Each s In types
            If Dir(fld & "\*." & s) <> "" Then
                HasFileType = True
                Exit Function
            End If
        Next s
    Next f
End Function

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

