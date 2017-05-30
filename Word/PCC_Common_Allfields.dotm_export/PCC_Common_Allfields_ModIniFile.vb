VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module      : modINIFile
' Description : Code to work with Windows Initialization (INI) files
' Source      : Total VB SourceBook 5
'
' Declarations for Windows INI File Manipulation procedures
Private Declare Function GetPrivateProfileInt _
  Lib "kernel32" _
  Alias "GetPrivateProfileIntA" _
  (ByVal strSection As String, _
    ByVal strKeyName As String, _
    ByVal lngDefault As Long, _
    ByVal strFilename As String) _
As Long
  
Private Declare Function GetPrivateProfileString _
  Lib "kernel32" _
  Alias "GetPrivateProfileStringA" _
  (ByVal strSection As String, _
    ByVal strKeyName As String, _
    ByVal strDefault As String, _
    ByVal strReturned As String, _
    ByVal lngSize As Long, _
    ByVal strFilename As String) _
As Long
    
Private Declare Function WinHelp _
  Lib "USER32" _
  Alias "WinHelpA" _
  (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) _
As Long

Private Declare Function WritePrivateProfileString _
  Lib "kernel32" _
  Alias "WritePrivateProfileStringA" _
  (ByVal strSection As String, _
    ByVal strKeyNam As String, _
    ByVal strValue As String, _
    ByVal strFilename As String) _
As Long

' API Calls to work with WIN.INI
Private Declare Function GetProfileInt _
  Lib "kernel32" _
  Alias "GetProfileIntA" _
  (ByVal strSection As String, _
    ByVal strKeyName As String, _
    ByVal lngDefault As Long) _
As Integer
  
Private Declare Function GetProfileString _
  Lib "kernel32" _
  Alias "GetProfileStringA" _
  (ByVal strSection As String, _
    ByVal strKeyName As String, _
    ByVal strDefault As String, _
    ByVal strReturned As String, ByVal intSize As Long) _
As Long
    
Private Declare Function WriteProfileString _
  Lib "kernel32" _
  Alias "WriteProfileStringA" _
  (ByVal strSection As String, _
    ByVal strKeyName As String, _
    ByVal strValue As String) _
As Integer

Public Function INIGetSettingInteger( _
  strSection As String, _
  strKeyName As String, _
  strFile As String) _
  As Integer
  ' Comments  : Returns an integer value from the specified INI file
  ' Parameters: strSection - Name of the section to look in
  '             strKeyName - Name of the key to look for
  '             strFile - Path and name of the INI file to look in
  ' Returns   : Integer value
  ' Source    : Total VB SourceBook 5
  '
  Dim intValue As Integer

  On Error GoTo PROC_ERR
  
  intValue = GetPrivateProfileInt(strSection, strKeyName, 0, strFile)

  INIGetSettingInteger = intValue

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIGetSettingInteger"
  Resume PROC_EXIT
  
End Function

Public Function INIGetSettingString( _
  strSection As String, _
  strKeyName As String, _
  strFile As String) _
  As String
  ' Comments  : Returns a string value from the specified INI file
  ' Parameters: strSection - Name of the section to look in
  '             strKeyName - Name of the key to look for
  '             strFile - Path and name of the INI file to look in
  ' Returns   : String value
  ' Source    : Total VB SourceBook 5
  '
  Dim strBuffer As String * 256
  Dim intSize As Integer

  On Error GoTo PROC_ERR
  
  intSize = GetPrivateProfileString(strSection, strKeyName, "", strBuffer, 256, strFile)

  INIGetSettingString = Left$(strBuffer, intSize)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIGetSettingString"
  Resume PROC_EXIT
  
End Function

Public Function INIWriteSetting( _
  strSection As String, _
  strKeyName As String, _
  strValue As String, _
  strFile As String) _
  As Integer
  ' Comments  : Writes the specified value to the specified INI file
  ' Parameters: strSection - section to write into
  '             strKeyName - key to write into
  '             strValue - value to write
  '             strFile - path and name of the INI file to write to
  ' Returns   : True if successful, False otherwise
  ' Source    : Total VB SourceBook 5
  '
  Dim intStatus As Integer

  On Error GoTo PROC_ERR
  
  intStatus = WritePrivateProfileString( _
    strSection, _
    strKeyName, _
    strValue, _
    strFile)

  INIWriteSetting = (intStatus <> 0)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIWriteSetting"
  Resume PROC_EXIT
  
End Function

Public Function WinINIGetSettingInteger( _
  strSection As String, _
  strKeyName As String) _
  As Integer
  ' Comments  : Returns an integer value from WIN.INI
  ' Parameters: strSection - name of the section to look in
  '             strKeyName - name of the key to look for
  ' Returns   : integer value
  ' Source    : Total VB SourceBook 5
  '
  Dim intValue As Integer

  On Error GoTo PROC_ERR
  
  intValue = GetProfileInt(strSection, strKeyName, 0)

  WinINIGetSettingInteger = intValue

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "WinINIGetSettingInteger"
  Resume PROC_EXIT
  
End Function

Public Function WinINIGetSettingString( _
  strSection As String, _
  strKeyName As String) As String
  ' Comments  : Returns a string value from the WIN.INI file
  ' Parameters: strSection - name of the section to look in
  '             strKeyName - name of the key to look for
  ' Returns   : string value
  ' Source    : Total VB SourceBook 5
  '
  Dim strBuffer As String * 256
  Dim intSize As Integer

  intSize = GetProfileString(strSection, strKeyName, "", strBuffer, 256)

  WinINIGetSettingString = Left$(strBuffer, intSize)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "WinINIGetSettingString"
  Resume PROC_EXIT
  
End Function

Public Function WinINIWriteSetting( _
  strSection As String, _
  strKeyName As String, _
  strValue As String) _
  As Integer
  ' Comments  : Writes the specified value to WIN.INI
  ' Parameters: strSection - section to write into
  '             strKeyName - key to write into
  '             strValue - value to write
  ' Returns   : True if successful, False otherwise
  ' Source    : Total VB SourceBook 5
  '
  Dim intStatus As Integer

  On Error GoTo PROC_ERR
  
  intStatus = WriteProfileString(strSection, strKeyName, strValue)

  WinINIWriteSetting = (intStatus <> 0)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "WinINIWriteSetting"
  Resume PROC_EXIT
  
End Function





