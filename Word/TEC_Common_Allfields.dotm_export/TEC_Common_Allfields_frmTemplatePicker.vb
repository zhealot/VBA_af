VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplatePicker 
   Caption         =   "TAS Template Picker"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   OleObjectBlob   =   "TEC_Common_Allfields_frmTemplatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' These templates have been prepared and developed for the TAS
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     info@allfields.co.nz, 04 978 7101
' Date:             February 2018
' Description:      Form used for picking template to load. Scans the
'                   Workgroup Templates folder for templates
'-----------------------------------------------------------------------------
Option Explicit



Private Sub lbGroup_Click()
    Dim fld As String
    fld = strWorkgroupTemplatesPath & "\" & lbGroup.Text
    'always show root folder content for "TAS"
    If lbGroup.Text = "TAS" Then
        fld = strWorkgroupTemplatesPath
    End If
    If Dir(fld, vbDirectory) = "" Then
       ThrowFatalError "This doesn't seem to be a template folder" & vbCr & vbCr & "Looking in " & """" & fld & """"
    Else
        'lbGroupName.Caption = "Business Group: " & lbGroup.Text
        Dim afn As Variant
        afn = GetFileList(fld & "\*." & ext)
        lbxWord.Clear
        lbxPPT.Clear
        lbxExcel.Clear
        If HasFileType(fld, cTypes) Then
            Dim sFile As String
            For Each sStr In cTypes
                sFile = Dir(fld & "\*." & sStr)
                Do While sFile <> ""
                    If Right(sFile, Len(sStr)) = sStr Then
                        If LCase(Left(sStr, 1)) = "d" Then
                            lbxWord.AddItem sFile
                        ElseIf LCase(Left(sStr, 1)) = "p" Then
                            lbxPPT.AddItem sFile
                        ElseIf LCase(Left(sStr, 1)) = "x" Then
                            lbxExcel.AddItem sFile
                        End If
                    End If
                    sFile = Dir() 'next file
                Loop
            Next sStr
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim iCounter As Integer
    Call ClearControls
    If Dir(strWorkgroupTemplatesPath, vbDirectory) = "" Then
        ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
        & vbCr & vbCr & "Looking in " & """" & strWorkgroupTemplatesPath & """"
    Else
        Dim objFld As Object
        Set objFld = ListSubFolders(strWorkgroupTemplatesPath)
        lbGroup.Clear
        If objFld.subfolders.Count > 0 Then
            Dim f
            For Each f In objFld.subfolders
                lbGroup.AddItem f.Name
            Next f
        End If
            Dim i As Integer
        For i = 0 To lbGroup.ListCount - 1
            If lbGroup.List(i) = DEFAULT_FOLDER Then
                lbGroup.ListIndex = i
                Exit For
            End If
        Next i
        'list files in folder root
        lbxWord.Clear
        lbxPPT.Clear
        lbxExcel.Clear
        If HasFileType(strWorkgroupTemplatesPath & "\" & DEFAULT_FOLDER, cTypes) Then
            Dim sFile As String
            For Each sStr In cTypes
                sFile = Dir(strWorkgroupTemplatesPath & "\" & DEFAULT_FOLDER & "\*." & sStr)
                Do While sFile <> ""

                    If Right(sFile, Len(sStr)) = sStr Then
                        Select Case LCase(Left(sStr, 1))
                        Case "d"    'Word
                            lbxWord.AddItem sFile
                        Case "p"    'PowerPoint
                            lbxPPT.AddItem sFile
                        Case "x"    'Excel
                            lbxExcel.AddItem sFile
                        End Select
                    End If
                    sFile = Dir() 'next file
                Loop
            Next sStr
        End If
    End If


        
    'Me.Show
End Sub
Private Sub lbxWord_Click()
    Dim i As Integer
    lbxPPT.ListIndex = -1
    lbxExcel.ListIndex = -1
    
    imgPreview.Picture = LoadPicture
    Dim spath As String
    spath = strWorkgroupTemplatesPath & "\" & frmTemplatePicker.lbGroup.List(lbGroup.ListIndex) & "\" & lbxWord.Text
    spath = Left(spath, InStrRev(spath, ".")) & imgEx
    If Not Dir(spath) = "" Then
        imgPreview.Picture = LoadPicture(spath, imgPreview.Width, imgPreview.Height)
        imgPreview.PictureSizeMode = fmPictureSizeModeZoom
    End If
End Sub

Private Sub lbxPPT_Click()
    Dim i As Integer
    lbxWord.ListIndex = -1
    lbxExcel.ListIndex = -1
    
    imgPreview.Picture = LoadPicture
    Dim spath As String
    spath = strWorkgroupTemplatesPath & "\" & frmTemplatePicker.lbGroup.List(lbGroup.ListIndex) & "\" & lbxPPT.Text
    spath = Left(spath, InStrRev(spath, ".")) & imgEx
    If Not Dir(spath) = "" Then
        imgPreview.Picture = LoadPicture(spath, imgPreview.Width, imgPreview.Height)
        imgPreview.PictureSizeMode = fmPictureSizeModeZoom
    End If
End Sub

Private Sub lbxExcel_Click()
    Dim i As Interior
    lbxWord.ListIndex = -1
    lbxPPT.ListIndex = -1
    
    imgPreview.Picture = LoadPicture
    Dim spath As String
    spath = strWorkgroupTemplatesPath & "\" & frmTemplatePicker.lbGroup.List(lbGroup.ListIndex) & "\" & lbxExcel.Text
    spath = Left(spath, InStrRev(spath, ".")) & imgEx
    If Not Dir(spath) = "" Then
        imgPreview.Picture = LoadPicture(spath, imgPreview.Width, imgPreview.Height)
        imgPreview.PictureSizeMode = fmPictureSizeModeZoom
    End If
End Sub

Sub cmbCancel_Click()
    Me.Hide
End Sub

Sub cmbOK_Click()
    Dim i As Integer
    Dim spath As String
'    Dim ob As OptionButton
    Dim found As Boolean
    found = False
    
    If lbxWord.ListIndex >= 0 Or lbxPPT.ListIndex >= 0 Or lbxExcel.ListIndex >= 0 Then
        spath = strWorkgroupTemplatesPath & "\" & lbGroup.List(lbGroup.ListIndex) & "\" & IIf(lbxWord.ListIndex > -1, lbxWord.Value, IIf(lbxPPT.ListIndex > -1, lbxPPT.Value, lbxExcel.Value))
        found = True
    Else
        found = False
    End If
    If found Then
        '###Unload Me
        Me.Hide
        'Create new document
        '###TODO, ability to launch PowerPoint or Excel app and add new file/presentation
        Dim sFileType As String
        sFileType = Trim(LCase(Right(spath, Len(spath) - InStrRev(spath, "."))))
        Select Case Left(sFileType, 2)
            Case "do"
                Dim newDoc As Document
                Set newDoc = Documents.Add(Template:=spath)
                Application.Visible = True
                newDoc.Activate
                newDoc.Windows(1).Activate
            Case "xl"
                Dim xApp As Excel.Application
                Dim xWb As Excel.Workbook
                Set xApp = New Excel.Application
                xApp.Visible = True
                Set xWb = xApp.Workbooks.Add(spath)
                xWb.Windows.Item(1).WindowState = xlMaximized
                xWb.Windows.Item(1).Activate
            Case "pp", "po"
                Dim pApp As PowerPoint.Application
                Set pApp = New PowerPoint.Application
                pApp.Presentations.Open spath, , msoCTrue
                pApp.Visible = msoTrue
                pApp.Presentations(1).Windows(1).Activate
                pApp.Activate
            Case Else
            End Select
    Else
        MsgBox "Please select a template from the list", vbOKOnly + vbCritical, "No template selected"
    End If
End Sub

Function Max(iNum1 As Integer, iNum2 As Integer) As Integer
    If iNum1 > iNum2 Then
        Max = iNum1
    Else
        Max = iNum2
    End If
End Function


Sub ClearControls()
    Me.lbxWord.Clear
    Me.lbxPPT.Clear
    Me.lbxExcel.Clear
End Sub

Function ListSubFolders(fld As String) As Object
'return 1 level subfolers in folder
    Dim FileSystem As Object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set ListSubFolders = FileSystem.GetFolder(fld)
End Function

