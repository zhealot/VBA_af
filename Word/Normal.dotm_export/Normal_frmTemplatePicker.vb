VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplatePicker 
   Caption         =   "TEC Templates"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   OleObjectBlob   =   "Normal_frmTemplatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'-----------------------------------------------------------------------------
' These templates have been prepared and developed for the TED
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     info@allfields.co.nz, 04 978 7101
' Date:             March 2011
' Description:      Form used for picking template to load. Scans the
'                   Workgroup Templates folder for templates with a .dotm
'                   extension, groups them by the text before the dash, and
'                   named by the text after the dash, less the extension
'-----------------------------------------------------------------------------
Option Explicit

'Double-clicking the top listbox is same as OK button
Private Sub lstStandard_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmbOK_Click
End Sub

Private Sub UserForm_Initialize()
    Dim iCounter As Integer
    '###test purpose
    'strWorkgroupTemplatesPath = "C:\Users\tao\Box Sync\1. Clients\TEC\TEC Templates provided"
    'initialize controls
    Call ClearControls
    Me.lb1st.Caption = ""
    Me.lb2nd.Caption = ""
    Me.lb3rd.Caption = ""
    If Dir(strWorkgroupTemplatesPath, vbDirectory) = "" Then
        ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
        & vbCr & vbCr & "Looking in " & """" & strWorkgroupTemplatesPath & """"
    Else
        Dim objFld As Object
        Set objFld = ListSubFolders(strWorkgroupTemplatesPath)
        If objFld.SubFolders.Count > 0 Then
            lbx1st.Clear
            Dim f
            For Each f In objFld.SubFolders
                If HasFileType(f.Path, ext) Then
                    lbx1st.AddItem f.Name
                End If
            Next f
        Else
            ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
                & vbCr & vbCr & "Looking in " & """" & strWorkgroupTemplatesPath & """"
        End If
        lb1st.Caption = Right(strWorkgroupTemplatesPath, Len(strWorkgroupTemplatesPath) - InStrRev(strWorkgroupTemplatesPath, "\"))
    End If
    'Me.Show
End Sub

Private Sub lbx1st_Click()
    Dim fld As String
    fld = strWorkgroupTemplatesPath & "\" & lbx1st.Text
    If Dir(fld, vbDirectory) = "" Then
        ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
        & vbCr & vbCr & "Looking in " & """" & fld & """"
    Else
        Call ClearControls
        Dim objFld As Object
        Set objFld = ListSubFolders(fld)
        If objFld.SubFolders.Count > 0 Or HasFileType(fld, ext) Then
            lbx2nd.Clear
            If objFld.SubFolders.Count > 0 Then
                Dim f
                For Each f In objFld.SubFolders
                    If HasFileType(f.Path, ext) Then
                        lbx2nd.AddItem f.Name
                    End If
                Next f
            End If
            If HasFileType(fld, ext) Then
                Dim sFile As String
                sFile = Dir(fld & "\*." & ext)
                Do While sFile <> ""
                    lbx3rd.AddItem sFile
                    sFile = Dir() 'next file
                Loop
            End If
        Else
            ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
                & vbCr & vbCr & "Looking in " & """" & strWorkgroupTemplatesPath & """"
        End If
        lb2nd.Caption = Right(fld, Len(fld) - InStrRev(fld, "\"))
    End If
End Sub

Private Sub lbx2nd_Click()
    Dim fld As String
    fld = strWorkgroupTemplatesPath & "\" & lbx1st.Text & "\" & lbx2nd.Text
    If Dir(fld, vbDirectory) = "" Then
        ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
        & vbCr & vbCr & "Looking in " & """" & fld & """"
    Else
        Call ClearControls
        Dim aFN As Variant
        aFN = GetFileList(fld & "\*." & ext)
        lb3rd.Caption = Right(fld, Len(fld) - InStrRev(fld, "\"))
        Dim fl
        Dim iBtn As Integer
        iBtn = 1
        For Each fl In aFN
            Dim str As String
            str = Left(lbx1st, 2)
            str = str + IIf(InStr(fl, " ") > 0, Left(fl, InStr(fl, " ")), Left(fl, InStr(fl, ".") - 1))
            If ColourValue(str) = 0 Then
                lbx3rd.AddItem fl
            Else
                Dim ob As OptionButton
                Set ob = Me.Controls("ob" & iBtn)
                ob.BackColor = ColourValue(str)
                ob.Caption = fl
                ob.Enabled = True
                iBtn = iBtn + 1
            End If
        Next fl
    End If
End Sub


Private Sub lbx3rd_Click()
    Dim i As Integer
    Dim ob As OptionButton
    If lbx3rd.ListIndex >= 0 Then
        For i = 1 To 5 Step 1
            Set ob = Me.Controls("ob" & i)
            ob.Value = False
        Next i
    End If
    imgPreview.Picture = LoadPicture
    Dim sPath As String
    sPath = imgPath & "\" & lbx1st.Text & "\" & lbx2nd.Text & "\" & lbx3rd.Text
    sPath = Left(sPath, InStrRev(sPath, ".")) & imgEx
    If Not Dir(sPath) = "" Then
        imgPreview.Picture = LoadPicture(sPath, imgPreview.Width, imgPreview.Height)
        imgPreview.PictureSizeMode = fmPictureSizeModeZoom
    End If
End Sub

Private Sub ob1_Click()
    ClearLv3 ob1.Caption
End Sub

Private Sub ob2_Click()
    ClearLv3 ob2.Caption
End Sub

Private Sub ob3_Click()
    ClearLv3 ob3.Caption
End Sub

Private Sub ob4_Click()
    ClearLv3 ob4.Caption
End Sub

Private Sub ob5_Click()
    ClearLv3 ob5.Caption
End Sub

'------------------------------------------------------------
'When the "Open Existing Document" is clicked, Word's Open
'dialog box is displayed. If it is OK'ed, the Menu is unloaded
'------------------------------------------------------------
Private Sub cmbOpen_Click()
    If Dialogs(wdDialogFileOpen).Show = -1 Then
        Unload frmTemplatePicker
    End If
End Sub

Sub cmbCancel_Click()
    Unload Me
    End
End Sub

Sub cmbOK_Click()
    Dim i As Integer
    Dim sPath As String
    Dim ob As OptionButton
    Dim found As Boolean
    found = False
    
    If lbx3rd.ListIndex >= 0 Then
        sPath = strWorkgroupTemplatesPath & "\" & lbx1st.Text & "\" & lbx2nd.Text & "\" & lbx3rd.Value
        found = True
    Else
        For i = 1 To 5 Step 1
            Set ob = Controls("ob" & i)
            If ob.Value Then
                sPath = strWorkgroupTemplatesPath & "\" & lbx1st.Text & "\" & lbx2nd.Text & "\" & ob.Caption
                found = True
                Exit For
            End If
        Next i
    End If
    If found Then
        '###Unload Me
        Me.Hide
        'Create new document
        Documents.Add Template:=sPath
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
    Dim ob As OptionButton
    Dim i As Integer
    For i = 1 To 5 Step 1
        Set ob = Me.Controls("ob" & i)
        ob.BackColor = 15790320
        ob.Caption = ""
        ob.Enabled = False
        ob.Value = False
    Next i
    Me.lbx3rd.Clear
End Sub

Function ClearLv3(cpt As String)
    Me.lbx3rd.ListIndex = -1
    imgPreview.Picture = LoadPicture
    Dim sPath As String
    sPath = strWorkgroupTemplatesPath & "\" & lbx1st.Text & "\" & lbx2nd.Text & "\" & cpt
    sPath = Left(sPath, InStrRev(sPath, ".")) & imgEx
    If Not Dir(sPath) = "" Then
        imgPreview.Picture = LoadPicture(sPath, imgPreview.Width, imgPreview.Height)
        imgPreview.PictureSizeMode = fmPictureSizeModeZoom
    End If
End Function

Sub test()
    Dim cc As ContentControl
    Dim iPD As IPictureDisp
    cc.Range.ShapeRange.Item (1)
    
End Sub
