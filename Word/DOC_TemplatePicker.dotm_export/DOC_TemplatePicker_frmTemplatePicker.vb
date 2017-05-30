VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplatePicker 
   Caption         =   "Template Picker"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   OleObjectBlob   =   "DOC_TemplatePicker_frmTemplatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' These templates have been prepared and developed for the Department Of Conservation
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             April 2017
' Description:      Form used for picking template to load. Scans the
'                   Workgroup Templates folder for templates with a .dotx
'                   extension, list them in alphabetical order, less the extension
'-----------------------------------------------------------------------------
Option Explicit
Dim OBArray() As New OBHandler    'array to hold OptionButton objects

Private Sub UserForm_Initialize()
    Frame1.Controls.Clear
    'strTemplatesPath = "D:\temp\MVCOT\workgroup\templates" '###
    If Dir(strTemplatesPath, vbDirectory) = "" Then
        ThrowFatalError "There doesnt seem to be any templates in the Workgroup Templates folder" _
        & vbCr & vbCr & "Looking in " & """" & strTemplatesPath & """"
        Exit Sub
    End If
    'check whether templates exist
    Dim sFileName As String
    sFileName = Dir(strTemplatesPath & "\*." & ext)
    If sFileName = "" Then
        MsgBox "No template found, please contact IT support"
        End
    End If
    'constants to place controls on user form
    Dim iOBCount As Integer
    Const TOP_BASE = 0          'optionbutton top base position
    Const LEFT_BASE = 20        'optionbutton left base position
    Const OB_OFFSET = 20        'optionbutton top offset
    Const OB_FONT = 12          'optionbutton font size
    Const OB_WIDTH = 300        'optionbutton width
    Const OB_HEIGHT = 20        'optionbutton height
    Const LABEL_OFFSET = 20     'label base offset
    Const LABEL_LEFT = 5        'label left base position
    Const LABEL_FONT = 14       'label font size
    Const LABEL_BOLD = True     'lable font bold
    
    Dim dblHeight As Double: dblHeight = 0 'current object top position
    Dim aGroup() As String
    Dim iGroup As Integer: iGroup = 0
    Dim aOtherGroup() As String
    Dim iOther As Integer: iOther = 0
    
    Do While sFileName <> ""
        If InStr(sFileName, "-") > 0 Then
            ReDim Preserve aGroup(iGroup)
            aGroup(iGroup) = sFileName
            iGroup = iGroup + 1
        Else
            If InStr(LCase(sFileName), "normal") <> 1 Then  'exclude 'Normal' template
                ReDim Preserve aOtherGroup(iOther)
                aOtherGroup(iOther) = sFileName
                iOther = iOther + 1
            End If
        End If
        sFileName = Dir()
    Loop
    
    Dim sLabel As String
    Dim CatItem() As String
    Dim iii As Integer
    Dim iOBHandler As Integer
    'add templates belong to group title
    If iGroup > 0 Then
        For iii = 0 To UBound(aGroup)
            If InStr(aGroup(iii), "-") > 0 Then
                CatItem = Split(aGroup(iii), "-")
                If sLabel <> CatItem(0) Then
                    sLabel = CatItem(0)
                    Debug.Print sLabel
                    Dim lb As MSForms.Label
                    Set lb = Frame1.Controls.Add("Forms.Label.1", "L1")
                    With lb
                        .Top = dblHeight
                        .Left = LABEL_LEFT
                        .Caption = sLabel
                        .Font.Size = LABEL_FONT
                        .Font.Name = "Segoe UI"
                        .Font.Bold = LABEL_BOLD
                    End With
                    dblHeight = dblHeight + LABEL_OFFSET
                End If
                Dim ob As MSForms.OptionButton
                Set ob = Frame1.Controls.Add("Forms.OptionButton.1", "o1")
                
                With ob
                    .Tag = CatItem(0)
                    .Top = dblHeight
                    .Left = LEFT_BASE
                    .Width = OB_WIDTH
                    .Height = OB_HEIGHT
                    .Font.Name = "Segoe UI"
                    .Font.Size = OB_FONT
                    .Caption = Left(CatItem(1), Len(CatItem(1)) - Len(ext) - 1)
                End With
                ReDim Preserve OBArray(iOBHandler)
                Set OBArray(iOBHandler).cb = ob
                OBArray(iOBHandler).Caption = ob.Caption
                OBArray(iOBHandler).Group = CatItem(0)
                iOBHandler = iOBHandler + 1
                dblHeight = dblHeight + OB_OFFSET
            End If
        Next
    End If
    
    'add templates that have no group
    If iOther > 0 Then
        Set lb = Frame1.Controls.Add("Forms.Label.1", "L1")
        With lb
            .Top = dblHeight
            .Left = LABEL_LEFT
            .Caption = "Other"
            .Font.Size = LABEL_FONT
            .Font.Name = "Segoe UI"
            .Font.Bold = LABEL_BOLD
        End With
        dblHeight = dblHeight + LABEL_OFFSET
        
        For iii = 0 To UBound(aOtherGroup)
            If aOtherGroup(iii) <> "" Then
                Set ob = Frame1.Controls.Add("Forms.OptionButton.1", "o1")
                With ob
                    .Tag = ""
                    .Top = dblHeight
                    .Left = LEFT_BASE
                    .Width = OB_WIDTH
                    .Height = OB_HEIGHT
                    .Font.Name = "Segoe UI"
                    .Font.Size = OB_FONT
                    .Caption = Left(aOtherGroup(iii), Len(aOtherGroup(iii)) - Len(ext) - 1)
                End With
                ReDim Preserve OBArray(iOBHandler)
                Set OBArray(iOBHandler).cb = ob
                OBArray(iOBHandler).Caption = ob.Caption
                OBArray(iOBHandler).Group = ""
                iOBHandler = iOBHandler + 1
                dblHeight = dblHeight + OB_OFFSET
            End If
        Next
    End If
    'set frame scroll bar to work
    With Me.Frame1
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = dblHeight + 20
    End With
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
    Me.Hide
    End
End Sub

Sub cmbOK_Click()
    Dim i As Integer
    Dim sPath As String
    Dim found As Boolean
    found = False
    Dim sFN As String
    Dim ob As Variant
    For Each ob In Frame1.Controls
        If TypeName(ob) = "OptionButton" Then
            If ob.Value Then
                sFN = ob.Tag & IIf(Len(ob.Tag) > 0, "-", "") & ob.Caption & "." & ext
                found = True
                Exit For
            End If
        End If
    Next ob
    If Not found Then
        MsgBox "Please select a template to open"
        Exit Sub
    End If
    Dim newDoc As Document
    sPath = strTemplatesPath & "\" & sFN
    Set newDoc = Documents.Add(Template:=sPath)
    Application.Visible = True
    frmTemplatePicker.Hide
    newDoc.Activate
End Sub
