Attribute VB_Name = "Allfields"
Public strActiveUser As String

'**************************************
' Replaces the contents of a bookmark
'**************************************
Public Sub ReplaceBookmarkText(sBookmark As String, sText As String, Optional Suppress As Boolean = True)
    Dim StartPos
    Dim EndPos
    
    If ActiveDocument.Bookmarks.Exists(sBookmark) Then
        Dim BMRange As Range
        Set BMRange = ActiveDocument.Bookmarks(sBookmark).Range
        BMRange.Text = sText
        ActiveDocument.Bookmarks.Add sBookmark, BMRange
    ElseIf Not Suppress Then
        MsgBox "Bookmark does not exist", vbCritical + vbOKOnly
        Exit Sub
    End If
    
End Sub

Public Function GetBookmarkText(sBookmark As String, Optional Suppress As Boolean = True) As String
    
    If ActiveDocument.Bookmarks.Exists(sBookmark) Then
        Dim BMRange As Range
        Set BMRange = ActiveDocument.Bookmarks(sBookmark).Range
        GetBookmarkText = BMRange.Text
    ElseIf Not Suppress Then
        MsgBox "Bookmark does not exist", vbCritical + vbOKOnly
        Exit Function
    End If
    
End Function


Function FileExists(fname) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(fname)
End Function

Function FolderExists(fname) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(fname)
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

Sub SetActiveUser(strUser As String)
    If UserExists(strUser) Then _
        strActiveUser = strUser
End Sub

Function GetActiveUser()
    If strActiveUser <> "" And UserExists(strActiveUser) Then
        GetActiveUser = strActiveUser
    Else
        GetActiveUser = GetDefaultUser
    End If
End Function

Function GetDefaultUser() As String
    Dim aUsers
    aUsers = IniOP.LoadIniSectionsArray(strUserIni)
    
    If UBound(aUsers) < 0 Then _
        ThrowFatalError "There are no User Profiles configured. Cannot retrieve default user."

    GetDefaultUser = aUsers(LBound(aUsers))
End Function

Function GetUserSetting(strProperty As String, strUser As String) As String
    GetUserSetting = GetINISetting(strUserIni, strUser, strProperty)
End Function

Function WriteUserSetting(strProperty As String, strValue As String, strUser As String) As Boolean
    WriteUserSetting = WriteINISetting(strUserIni, strUser, strProperty, strValue)
End Function

Function GetINISetting(strINI As String, strGroup As String, strProperty As String) As String
    If FileExists(strINI) Then
        GetINISetting = System.PrivateProfileString(strINI, strGroup, strProperty)
    Else
        ThrowFatalError "The requested INI file " & strINI & " could not be found"
    End If
End Function

Function WriteINISetting(strINI As String, strGroup As String, strProperty As String, strValue As String) As Boolean
    Dim i As Integer
    If GetINISetting(strINI, strGroup, strProperty) = strValue Then GoTo Write_Done
    If FileExists(strINI) Then
        i = 1
        On Error GoTo Write_Failed
        System.PrivateProfileString(strINI, strGroup, strProperty) = strValue

        GoTo Write_Done
    Else
        ThrowFatalError "The requested INI file " & strINI & " could not be found"
    End If

Write_Done:
    WriteINISetting = True
    Exit Function

Write_Failed:
    If i <= MAX_INI_WRITE_ATTEMPTS Then
        i = i + 1
        Resume
    Else
        MsgBox "Writing to INI file failed after " & MAX_INI_WRITE_ATTEMPTS & _
                " attempts. Please alert IT." & vbCr & vbCr & "File: " & _
                    strINI, vbCritical + vbOKOnly, "Fatal: Write Error"
        
        WriteINISetting = False
        Exit Function
    End If
    
End Function

Function UserExists(strUser As String) As Boolean
    UserExists = IniOP.CheckIfIniSectionExists(strUser, strUserIni)
End Function

Sub ReplaceWorkunitBranding(strUser As String)
    If Not UserExists(strUser) Then Exit Sub
    ' Added to get around the Objective Integration errors
    ' when using drawing objects
    Dim iErrCount As Integer
    iErrCount = 0
    'On Error GoTo ErrorHandler
    
    Dim sh As InlineShape
    Dim logoPath As String
    Dim strWorkunit As String
    Dim fRatio As Double
    strWorkunit = GetUserSetting("Workunit", strUser)
    
    If ActiveDocument.Bookmarks.Exists(BRANDING_LOGO_BKM) Then
        logoPath = GetWorkunitImage(strWorkunit)
        If logoPath <> "" Then
            ReplaceBookmarkText BRANDING_LOGO_BKM, ""
            Set sh = ActiveDocument.Bookmarks(BRANDING_LOGO_BKM).Range.InlineShapes.AddPicture(Filename:=logoPath)
            fRatio = sh.Width / sh.Height
            sh.Height = 90
            sh.Width = sh.Height * fRatio
        End If
    End If
    If ActiveDocument.Bookmarks.Exists(BRANDING_TEXT_BKM) Then _
        ReplaceBookmarkText BRANDING_TEXT_BKM, GetUserAddress(strUser)
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 4605 Then
        If iErrCount < 3 Then
            iErrCount = iErrCount + 1
            Resume
        Else
            MsgBox "Maximium Objective Integration error suppressions used." & _
             vbCr & vbCr & "Please contact IT"
        End If
    End If
End Sub

Function GetWorkunitImage(strWorkunit As String) As String
    Dim sh As Shape
    Dim logoPath As String
    logoPath = strWorkunitLogosPath & "\Core\" & strWorkunit & ".jpg"
    If Not FileExists(logoPath) Then
        ShowError ("Can not locate logo for " & strWorkunit & vbCr & vbCr & _
            "Looking in " & logoPath & vbCr & vbCr & _
            "Using MSD Logo Instead")
        Exit Function
    End If
    
    GetWorkunitImage = logoPath
End Function

Function GetUserAddress(strUser As String) As String

    Dim strStreet As String, strPostal As String, strCity As String, _
        strPhone As String, strFax As String, strAddress As String, _
        strWorkunit As String
    
    strWorkunit = GetUserSetting("Workunit", strUser)
    strStreet = GetUserSetting("BusinessStreetAddress", strUser)
    strPostal = GetUserSetting("BusinessPostalAddress", strUser)
    ' there is no special setting for this WU, use default
    If strAddress = "" And strPostal = "" Then
        strStreet = GetWorkunitAddressPart(strWorkunit, "Street")
        strPostal = GetWorkunitAddressPart(strWorkunit, "Postal")
        strCity = GetWorkunitAddressPart(strWorkunit, "City")
        strPhone = GetWorkunitAddressPart(strWorkunit, "Telephone")
        strFax = GetWorkunitAddressPart(strWorkunit, "Facsimile")
    Else
        strCity = GetUserSetting("BusinessCity", strUser)
        strPhone = GetUserSetting("BusinessPhone", strUser)
        strFax = GetUserSetting("BusinessFacsimile", strUser)
    End If
    
    strPhone = "Telephone: " & strPhone
    strFax = "Facsimile: " & strFax
    
    strAddress = strStreet
    If strStreet <> "" Then strAddress = strAddress & ", "
    strAddress = strAddress & strPostal & ", " & strCity
    
    GetUserAddress = strAddress & GetDivider() & strPhone & GetDivider() & strFax
    
End Function

Function GetWorkunitAddressPart(Workunit As String, Part As String) As String
    Dim Value As String
    Value = GetINISetting(strAddressConfigPath, Workunit, Part)
    If Value = "" Then _
        Value = GetINISetting(strAddressConfigPath, "Default", Part)
    
    GetWorkunitAddressPart = Value
End Function

Function GetDivider() As String
    GetDivider = " " & Chr(151) & " "
End Function

Public Function GetActiveWorkunitCode()
    GetActiveWorkunitCode = GetWorkunitCode(GetUserSetting("Workunit", GetActiveUser))
End Function

Public Function GetWorkunitCode(strWorkunit As String, Optional Default As String = "MSD") As String
    If strWorkunit = "" Then
        GetWorkunitCode = Default
        Exit Function
    End If
    
    Dim aWorkunits As Variant
    Dim iPos As Integer
    
    aWorkunits = GetINIGroupMatrix("Workunits")
    
    For iPos = LBound(aWorkunits) To UBound(aWorkunits)
        If aWorkunits(iPos, 2) = strWorkunit Then
            GetWorkunitCode = aWorkunits(iPos, 1)
            Exit Function
        End If
    Next iPos
    
    MsgBox "Cound not find Workunit code for " & strWorkunit & vbCr & vbCr & _
            "Assuming MSD for workunit code"
    GetWorkunitCode = "MSD"
    
End Function

Public Sub EnableToolbar(strToolbarName As String)
    
    On Error Resume Next
    With Application.CommandBars(strToolbarName)
        .Enabled = True
        .Visible = True
    End With
    On Error GoTo 0

End Sub

Public Sub DisableToolbar(strToolbarName As String)
    
    On Error Resume Next
    With Application.CommandBars(strToolbarName)
    '    .Enabled = False
        .Visible = False
    End With
    On Error GoTo 0

End Sub

Function GetMinisterMatrix()

    Dim Lines() As String
    Dim MinisterMatrix() As String
    Dim iLine As Integer, iMinisterPos As Integer, iPos As Integer
    Dim MIN_INX As Integer, PORT_INX As Integer
    Dim sParts() As String, sPortfolios() As String
    MIN_INX = 1
    PORT_INX = 2
    iMinisterPos = 0
    Lines = IniOP.LoadIniSectionArray("Ministers", strMSDGlobalPath)

    For iLine = LBound(Lines) To UBound(Lines)
        ReDim Preserve MinisterMatrix(MIN_INX To PORT_INX, 0 To iMinisterPos)
        sParts = Split(Lines(iLine), "=", 2)
        sPortfolios = Split(sParts(1), "#")
        MinisterMatrix(MIN_INX, iMinisterPos) = sParts(0)
        MinisterMatrix(PORT_INX, iMinisterPos) = sPortfolios(0)
        ' If the minister has more than one portfolio the add additional lines here
        If UBound(sPortfolios) > 0 Then
            For iPos = 1 To UBound(sPortfolios)
                iMinisterPos = iMinisterPos + 1
                ReDim Preserve MinisterMatrix(MIN_INX To PORT_INX, 0 To iMinisterPos)
                MinisterMatrix(MIN_INX, iMinisterPos) = sParts(0)
                MinisterMatrix(PORT_INX, iMinisterPos) = sPortfolios(iPos)
            Next iPos
        End If
        iMinisterPos = iMinisterPos + 1
        
    Next iLine
    
    GetMinisterMatrix = Transpose(MinisterMatrix)
    
End Function


Function GetINIGroupMatrix(strGroup As String, Optional strIniFile As String = "")
    
    If strIniFile = "" Then strIniFile = strMSDGlobalPath
    
    Dim GroupMatrix() As String
    Dim Lines() As String, sParts() As String
    Dim iLine As Integer, KEY_INX As Integer, VAL_INX As Integer
    KEY_INX = 1
    VAL_INX = 2
    Lines = IniOP.LoadIniSectionArray(strGroup, strMSDGlobalPath)
    ReDim GroupMatrix(KEY_INX To VAL_INX, LBound(Lines) To UBound(Lines))
    For iLine = LBound(Lines) To UBound(Lines)
        sParts = Split(Lines(iLine), "=", 2)
        GroupMatrix(KEY_INX, iLine) = sParts(0)
        GroupMatrix(VAL_INX, iLine) = sParts(1)
    Next iLine
    GetINIGroupMatrix = Transpose(GroupMatrix)
    
End Function

' Largely similar to the PopulateCtl function, this returns a two-dimensional array
' of ministers to their respective portfolios based on the INI file
Function GetINIGroupMatrix_Old(strGroup As String, Optional strIniFile As String = "") As Variant
    
    If strIniFile = "" Then strIniFile = strMSDGlobalPath
    
    Dim oFSO As New FileSystemObject
    Dim oFS As TextStream
    
    Dim sText As String, strGroupHeading As String
    
    Dim GroupMatrix() As Variant
    Dim iLinePos As Integer
    Dim MIN_INX As Integer, PORT_INX As Integer
    Dim aLine As Variant, aPortfolios As Variant
    Dim iPos As Integer
    
    iGroupPos = 0
    KEY_IDX = 1
    VAL_IDX = 2
    
    If FileExists(strIniFile) Then
        Set oFS = oFSO.OpenTextFile(strIniFile)

        Do Until oFS.AtEndOfStream
            sText = oFS.Read(1)
            If sText = "[" Then
                sText = oFS.ReadLine
                If sText = strGroup & "]" Then
                    ' Found the group heading, read the contents until another group happens
                    Do
                        sText = oFS.ReadLine
                        If Left(sText, 1) <> ";" And Trim(sText) <> "" And Left(sText, 1) <> "[" Then
                            ReDim Preserve GroupMatrix(KEY_IDX To VAL_IDX, 0 To iGroupPos)
                            sText = Trim(sText)
                            aLine = Split(sText, "=", 2)
                            GroupMatrix(KEY_IDX, iGroupPos) = aLine(0)
                            GroupMatrix(VAL_IDX, iGroupPos) = aLine(1)
                            iGroupPos = iGroupPos + 1
                        End If
                        
                    Loop Until oFS.AtEndOfStream Or Left(sText, 1) = "["
                    
                    GetINIGroupMatrix = Transpose(GroupMatrix)
                    
                    Exit Function
                End If
            Else
                oFS.SkipLine
            End If
        Loop
        oFS.Close
    End If

End Function

Function Transpose(StringArray() As String) As String()
    
    Dim iMin As Integer, iMax As Integer
    Dim jMin As Integer, jMax As Integer
    Dim iPos As Integer, jPos As Integer
    Dim Transposed() As String
    
    iMin = LBound(StringArray)
    iMax = UBound(StringArray)
    jMin = LBound(StringArray, 2)
    jMax = UBound(StringArray, 2)
    
    ReDim Transposed(jMin To jMax, iMin To iMax) As String
    
    For iPos = iMin To iMax
        For jPos = jMin To jMax
            Transposed(jPos, iPos) = StringArray(iPos, jPos)
        Next jPos
    Next iPos
    
    Transpose = Transposed
    
End Function
