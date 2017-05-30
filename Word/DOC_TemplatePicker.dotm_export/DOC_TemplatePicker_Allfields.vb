Attribute VB_Name = "Allfields"
Public TOOLKIT_LOADED As Boolean
Public Const ext = "dotx" '"docx,docm,docx,doc,dotx,dotm,dot,xlsx,xlsm,xlsb,xls,xltx,xltm,xlt,pptx,pptm,ppt,potx,potm,pot,ppsx,ppsm,pps"
Public Const imgEx = "jpg"
Public strTemplatesPath As String
Public imgPath As String

Public Sub setDefaultTab(control As IRibbonUI)

End Sub

Sub Autoexec()
    If TOOLKIT_LOADED Then Exit Sub
    On Error Resume Next
    Dim sPath As String
    
    sPath = Options.DefaultFilePath(wdWorkgroupTemplatesPath) & "\DOC Templates" 'setup default templates path
    strTemplatesPath = sPath & "\"
    strHelpPath = sPath 'set Help document path
    TOOLKIT_LOADED = True
    imgPath = sPath
End Sub


Public Sub ShowTemplatesMenu(control As IRibbonControl)
    Allfields.Autoexec
    CheckRequirements
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

'Function HasFileType(fld As String, types As Collection) As Boolean
''check if folder has file(s) of extension, include subfolers
'    HasFileType = False
'    'check folder root
'    If Not FolderExists(fld) Then Exit Function
'    Dim s As Variant
'    For Each s In types
'        If Dir(fld & "\*." & s) <> "" Then
'            HasFileType = True
'            Exit Function
'        End If
'    Next s
'
'    'check subfolder
'    Dim f
'    Dim objFld As Object
'    Set objFld = ListSubFolders(fld)
'    If objFld Is Nothing Then Exit Function
'    For Each f In objFld.SubFolders
'        For Each s In types
'            If Dir(fld & "\*." & s) <> "" Then
'                HasFileType = True
'                Exit Function
'            End If
'        Next s
'    Next f
'End Function

Sub test()
    
End Sub

