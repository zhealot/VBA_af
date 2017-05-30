VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherTemplatesPicker 
   Caption         =   "Other Ministry Templates"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   OleObjectBlob   =   "MVCOT_Common_frmOtherTemplatesPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtherTemplatesPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------
' These templates have been prepared and developed for the MED
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     info@allfields.co.nz, 04 978 7101
' Date:             February 2011
' Description:      Form used for picking template to load. Scans the
'                   Other Templates folder for templates with
'                   a .dot* extension. Makes use of a hidden list box to
'                   remember what extension each template has
'-----------------------------------------------------------------------------
Option Explicit

'This routine displays any sub folders for the selected folder
Private Sub cboBranchFolder_Change()
    Dim strSubFolder As String, strBranch As String
    Dim X As Integer, i As Integer
    Dim ListArray() As String
    
    cboSubFolder.Clear
    
    strSubFolder = strOtherTemplatesPath & "\" & cboBranchFolder.Text & "\"
    
    X = 0
    ReDim ListArray(X)
    
    strBranch = Dir(strSubFolder, vbDirectory)
    
    Do While strBranch <> ""
        If strBranch <> "." And strBranch <> ".." Then
            If (GetAttr(strSubFolder & strBranch) And vbDirectory) = vbDirectory Then
                ListArray(X) = strBranch
                X = X + 1
                ReDim Preserve ListArray(X)
            End If
        End If
        strBranch = Dir
    Loop

    Call TemplatePicker.QuickSort(ListArray, LBound(ListArray), UBound(ListArray))

    For i = 0 To X
        cboSubFolder.AddItem ListArray(i)
    Next
    If cboSubFolder.ListCount > 0 Then
        cboSubFolder.ListIndex = 0
    End If
    
    Erase ListArray()
    
    Call FillTemplates

End Sub

Private Sub cboSubFolder_Change()
    Call FillTemplates
End Sub

'This routine ensures that a selection is made, then creates a new document based on
'which template was selected.
Private Sub cmbOK_Click()
    Dim strFileTitle As String
    Dim strFileExtension As String
    Dim strNewFile As String
    Dim strFolder As String
    
    On Error GoTo ErrorHandler
    
    If lstTemplates.ListIndex = -1 Then
        MsgBox "You must select a Template!" & vbCr & "   (or select Cancel to exit)", _
            vbExclamation, "No Selection"
        Exit Sub
    Else
        If cboBranchFolder.Text <> "" Then
            strFolder = cboBranchFolder.Text & "\"
        Else
            strFolder = ""
        End If
        If cboSubFolder.Text <> "" Then
            strFolder = strFolder & cboSubFolder.Text & "\"
        End If
        
        strFileTitle = lstTemplates.List(lstTemplates.ListIndex)
        strFileExtension = lstExtensions.List(lstTemplates.ListIndex)
        
        strNewFile = strOtherTemplatesPath & "\" & strFolder & strFileTitle & "." & strFileExtension
        
        strFileExtension = LCase(strFileExtension)
        
        Unload Me
        'create new document based on selected template
        If Left(strFileExtension, 3) = "doc" Or Left(strFileExtension, 3) = "dot" Then
            Documents.Add strNewFile
        ElseIf Left(strFileExtension, 3) = "xls" Or Left(strFileExtension, 3) = "xlt" Then
            Dim oExcel, oWB
            Set oExcel = CreateObject("Excel.Application")
            oExcel.Visible = True
            Set oWB = oExcel.Workbooks.Open(strNewFile)
        ElseIf Left(strFileExtension, 3) = "ppt" Or Left(strFileExtension, 3) = "pot" Then
            Dim objPPT, objPresentation, strFile As String
            Set objPPT = CreateObject("PowerPoint.Application")
            objPPT.Visible = True
            Set objPresentation = objPPT.Presentations.Open(strNewFile, True)
        End If
        
    End If
    Exit Sub
    
    'error check primarily if network connection is lost
ErrorHandler:
        If Err.Number = 5137 Then
            MsgBox "That template cannot be found." & vbCr & vbCr & "Please check Network connections and try again", vbCritical, "File Not Found"
            Exit Sub
        Else
            MsgBox Err.Number & vbCr & vbCr & Err.Description
        End If
        Resume Next
End Sub

Private Sub lstTemplates_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmbOK_Click
End Sub

Private Sub UserForm_Initialize()
    Dim ListArray()  As String
    Dim i As Integer, X As Integer
    Dim strExtensionless As String, strTemplateFile As String
    Dim strBranchDir As String
    Dim strBranch As String
    Dim Drive As String
    
    On Error GoTo ErrorHandler
    
    strBranchDir = strOtherTemplatesPath & "\"
       
    X = 0
    ReDim ListArray(X)
    'Fill the Branch combobox
    strBranch = Dir(strBranchDir, vbDirectory)
    
    Do While strBranch <> ""
        If strBranch <> "." And strBranch <> ".." Then
           If (GetAttr(strBranchDir & strBranch) And vbDirectory) = vbDirectory Then
                cboBranchFolder.AddItem strBranch
            End If
        End If
        strBranch = Dir
    Loop

    If cboBranchFolder.ListCount > 0 Then
        cboBranchFolder.ListIndex = 0
    End If
    
    
    Erase ListArray()
    
    Exit Sub
    
    'if the Network drive is not available, a message is displayed
ErrorHandler:
        If Err.Number = 68 Then
            ThrowFatalError "A Network drive is unavailable"
        Else
            MsgBox Err.Number & vbCr & vbCr & Err.Description
        End If
        Resume Next

End Sub

Sub cmbCancel_Click()
    Unload Me
    End
End Sub

Sub FillTemplates()
    Dim strTemplateFile As String
    Dim strFolder As String
    Dim ListArray() As String
    Dim X As Integer, i As Integer
    Dim strExtensionless As String
    
    lstTemplates.Clear
    lstExtensions.Clear
    
    If cboBranchFolder.Text <> "" Then
        strFolder = cboBranchFolder & "\"
    Else
        strFolder = ""
    End If
    If cboSubFolder.Text <> "" Then
        strFolder = strFolder & cboSubFolder & "\"
    End If
    
    strTemplateFile = Dir(strOtherTemplatesPath & "\" & strFolder)
    
    X = 0
    ReDim Preserve ListArray(X)
    ' Start the loop.
    Do While strTemplateFile <> ""
        If IsAllowedExtension(Mid(strTemplateFile, InStrRev(strTemplateFile, ".") + 1)) Then
            ListArray(X) = strTemplateFile
            X = X + 1
            ReDim Preserve ListArray(X)
            'Call Dir again without arguments to return the next file in the same directory.
        End If
        strTemplateFile = Dir
    Loop
       
    Call TemplatePicker.QuickSort(ListArray, LBound(ListArray), UBound(ListArray))
    
    For i = 0 To X
        If (ListArray(i) <> "") Then
            lstTemplates.AddItem GetFileTitle(ListArray(i))
            lstExtensions.AddItem GetFileExtension(ListArray(i))
        End If
    Next
    
    
    
    Erase ListArray()
End Sub

Private Function IsAllowedExtension(strExtension As String)
    strExtension = LCase(strExtension)
    If Left(strExtension, 3) = "doc" Then
        IsAllowedExtension = True
    ElseIf Left(strExtension, 3) = "dot" Then
        IsAllowedExtension = True
    ElseIf Left(strExtension, 3) = "xls" Then
        IsAllowedExtension = True
    ElseIf Left(strExtension, 3) = "xlt" Then
        IsAllowedExtension = True
    ElseIf Left(strExtension, 3) = "ppt" Then
        IsAllowedExtension = True
    ElseIf Left(strExtension, 3) = "pot" Then
        IsAllowedExtension = True
    Else
        IsAllowedExtension = False
    End If
End Function

