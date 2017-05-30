VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Data Migration Tool"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7380
   OleObjectBlob   =   "DMTool_Special_UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnNext_Click()
    If UpdateMeta <> 0 Then
        'On Error Resume Next
        If InputSheet Is ActiveSheet And ActiveCell.Row < InputSheet.Rows.Count Then
            ActiveCell.Offset(1, 0).Select
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub btnPre_Click()
    If UpdateMeta <> 0 Then
        On Error Resume Next
        If ActiveSheet.Name = InputSheet.Name And ActiveCell.Row > 1 Then
            ActiveCell.Offset(-1, 0).Select
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub cbAction_Change()
    If EnEvents = False Then
        Exit Sub
    End If
    
    UpdateCell intISAction, cbAction
    If cbAction.Value = "Active" Then
        'cbPS.Enabled = True
        cbSC.Enabled = True
        UpdateComboBox UserForm1.cbSC, ws.UsedRange, intSC, RgSC
        cbSiteBranch.Enabled = True
        cbSitePro.Enabled = True
        cbSiteSC.Enabled = True
        cbSite1.Enabled = True
        cbSite2.Enabled = True
        cbSite3.Enabled = True
        cbLibrary.Enabled = True
        tbDocuSet.Enabled = True
        cbFunction.Enabled = True
        cbFunction2.Enabled = True
        cbFunctionEx.Enabled = True
        cbFunctionEx2.Enabled = True
        cbFunctionEx3.Enabled = True
    Else
        'cbPS.Enabled = False
        EnEvents = False
        DisCB cbSC, intISSC
        DisCB cbSiteBranch, intISSiteBranch
        DisCB cbSitePro, intISSitePro
        DisCB cbSiteSC, intISSiteSC
        DisCB cbSite1, intISSite1
        DisCB cbSite2, intISSite2
        DisCB cbSite3, intISSite3
        DisCB cbLibrary, intISLibrary
        DisCB cbFunction, intISFunction
        DisCB cbFunction2, intISFunction2
        DisCB cbFunctionEx, intISFunctionEx
        DisCB cbFunctionEx2, intISFunctionEx2
        DisCB cbFunctionEx3, intISFunctionEx3
        tbDocuSet.Value = ""
        tbDocuSet.Enabled = False
        EnEvents = True
    End If
End Sub

Private Sub cbSC_Change()
    If EnEvents = False Then
        Exit Sub
    End If
    'clear afterwards cells
    
    GreyCB cbSite1
    GreyCB cbSite2
    GreyCB cbSite3
    GreyCB cbLibrary
    GreyCB cbFunction
    GreyCB cbFunction2
    GreyCB cbFunctionEx
    GreyCB cbFunctionEx2
    GreyCB cbFunctionEx3
    
    Set RgSC = RangeOfValue(cbSC.Value, intSC, ws.UsedRange)
    UpdateComboBox cbSiteBranch, RgSC, intSiteBranch, RgSiteBranch
    If cbSiteBranch.ListCount = 0 Then
        GreyCB cbSiteBranch
        UpdateComboBox cbSitePro, RgSC, intSitePro, RgSitePro
        UpdateComboBox cbSiteSC, RgSC, intSiteSC, RgSiteSC
    Else
        cbSiteBranch.Enabled = True
        GreyCB cbSitePro
        GreyCB cbSiteSC
    End If
'    If cbSiteSC.ListCount = 0 Then
'        UpdateComboBox cbSite1, RgSC, intSite1, RgSite1
'    Else '###
'        If MinusRg(RgSC, RgSiteSC) Is Nothing Then
'            UpdateComboBox cbSite1, RgSC, intSite1, RgSite1
'        ElseIf Intersect(MinusRg(RgSC, RgSiteSC), ws.Columns(intSite1)) Is Nothing Then
'            UpdateComboBox cbSite1, RgSC, intSite1, RgSite1
'        Else
'            UpdateComboBox cbSite1, MinusRg(RgSC, RgSiteSC), intSite1, RgSite1
'        End If
'    End If
    
    UpdateCell intISSC, cbSC
    
End Sub


Private Sub cbSiteBranch_Change()
    If EnEvents = False Then
        Exit Sub
    End If
        
    Set RgSiteBranch = RangeOfValue(cbSiteBranch.Value, intSiteBranch, RgSC)
    cbSitePro.Enabled = True
    UpdateComboBox cbSitePro, RgSiteBranch, intSitePro, RgSitePro
    cbSiteSC.Enabled = True
    UpdateComboBox cbSiteSC, RgSiteBranch, intSiteSC, RgSiteSC
    If cbSiteSC.ListCount = 0 Then
        UpdateComboBox cbSite1, RgSiteBranch, intSite1, RgSite1
    End If
    
    UpdateCell intISSiteBranch, cbSiteBranch
End Sub

Private Sub cbSitePro_Change()
    If EnEvents = False Then
        Exit Sub
    End If

    Set RgSitePro = RangeOfValue(cbSitePro.Value, intSitePro, RgSiteBranch)
    UpdateComboBox cbSiteSC, RgSitePro, intSiteSC, RgSiteSC
    UpdateCell intISSitePro, cbSitePro
End Sub

Private Sub cbSiteSC_Change()
    If EnEvents = False Then
        Exit Sub
    End If
    
    Set RgSiteSC = RangeOfValue(cbSiteSC.Value, intSiteSC, RgSitePro)
    cbSite1.Enabled = True
    cbSite2.Enabled = True
    cbSite3.Enabled = True
    UpdateComboBox cbSite1, RgSiteSC, intSite1, RgSite1
    UpdateComboBox cbSite2, RgSiteSC, intSite2, RgSite2
    UpdateComboBox cbSite3, RgSiteSC, intSite3, RgSite3
    cbLibrary.Enabled = True
    If cbSite1.ListCount = 0 Then
        UpdateComboBox cbLibrary, RgSiteSC, intLibrary, RgLibrary
    Else
        '###
        If MinusRg(RgSiteSC, RgSite1) Is Nothing Then
            UpdateComboBox cbLibrary, RgSiteSC, intLibrary, RgLibrary
        ElseIf Intersect(MinusRg(RgSiteSC, RgSite1), ws.Columns(intLibrary)) Is Nothing Then
            UpdateComboBox cbLibrary, RgSiteSC, intLibrary, RgLibrary
        Else
            UpdateComboBox cbLibrary, MinusRg(RgSiteSC, RgSite1), intLibrary, RgLibrary
        End If
    End If
    'GreyCB cbTopic
    UpdateCell intISSiteSC, cbSiteSC
End Sub

Private Sub cbSite1_Change()
    If EnEvents = False Then
        Exit Sub
    End If
    
    If cbSiteSC.ListIndex = -1 Then
        Set RgSite1 = RangeOfValue(cbSite1.Value, intSite1, RgSC)
    Else
        Set RgSite1 = RangeOfValue(cbSite1.Value, intSite1, RgSiteSC)
    End If
    
    UpdateComboBox cbSite2, RgSite1, intSite2, RgSite2
    UpdateComboBox cbSite3, RgSite1, intSite3, RgSite3
    UpdateComboBox cbLibrary, RgSite1, intLibrary, RgLibrary
    
    UpdateCell intISSite1, cbSite1
End Sub

Private Sub cbSite2_Change()
    If EnEvents = False Then
        Exit Sub
    End If
    
    Set RgSite2 = RangeOfValue(cbSite2.Value, intSite2, RgSite1)
    UpdateComboBox cbSite3, RgSite2, intSite3, RgSite3
    UpdateComboBox cbLibrary, RgSite2, intLibrary, RgLibrary
    
    UpdateCell intISSite2, cbSite2
End Sub

Private Sub cbSite3_Change()
    If EnEvents = False Then Exit Sub
    
    Set RgSite3 = RangeOfValue(cbSite3.Value, intSite3, RgSite2)
    UpdateComboBox cbLibrary, RgSite3, intLibrary, RgLibrary
    
    UpdateCell intISSite3, cbSite3
End Sub

Private Sub cbLibrary_Change()
    If EnEvents = False Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim tbMetas As MSForms.TextBox
    Dim rg As Range
    
    Set RgLibrary = RangeOfValue(cbLibrary.Value, intLibrary, RgSite2)
    UpdateComboBox cbFunction, RgLibrary, intFunction, RgFunction
    'cbTopic.Enabled = IIf(cbTopic.ListCount > 0, True, False)
    cbFunction.Enabled = IIf(cbFunction.ListCount > 0, True, False)
    UpdateCell intISLibrary, cbLibrary
    
    'fill meta data
    Do While TBCollection.Count > 0
        UserForm1.Controls.Remove TBCollection.Item(1).Name
        TBCollection.Remove 1
    Loop
    UserForm1.Height = dbFormHeight
    
    'populate meta data area
    UpdateComboBox cbTemp, RgLibrary, intMeta, RgMeta
    If cbTemp.ListCount > 0 Then
        For i = 0 To cbTemp.ListCount - 1
            Set tbMetas = AddControl("Forms.TextBox.1", "tbMetaKey" & i, UserForm1.lbMeta.Top + i * 16, UserForm1.tbDocuSet.Left, 16, 100)
            TBCollection.Add tbMetas, tbMetas.Name
            tbMetas.Value = cbTemp.List(i)
            tbMetas.Font.Size = 8
            tbMetas.Enabled = False
            Set tbMetas = AddControl("Forms.TextBox.1", "tbMetaValue" & i, UserForm1.lbMeta.Top + i * 16, UserForm1.tbDocuSet.Left + 110, 16, 100)
            tbMetas.Font.Size = 8
            TBCollection.Add tbMetas, tbMetas.Name
            UserForm1.Height = UserForm1.Height + 16
        Next i
    End If
    
    Set rg = RgLibrary.Find(cbLibrary.Value, LookIn:=xlValues, lookat:=xlWhole)
    If Not rg Is Nothing Then
        If rg.Offset(0, 1).Value = "Yes" Then
            Me.tbDocuSet.Enabled = True
            Me.tbDocuSet.BackColor = &H80000005
        Else
            Me.tbDocuSet.Enabled = False
            Me.tbDocuSet.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub cbFunction_Change()
    If EnEvents = False Then Exit Sub
        
    Set RgFunction = RangeOfValue(cbFunction.Value, intFunction, RgLibrary)
    UpdateComboBox cbFunction2, RgFunction, intFunction2, RgFunction2
    UpdateCell intISFunction, cbFunction
End Sub

Private Sub cbFunction2_Change()
    If EnEvents = False Then Exit Sub
        
    Set RgFunction2 = RangeOfValue(cbFunction2.Value, intFunction2, RgFunction)
    UpdateComboBox cbFunctionEx, RgFunction2, intFunctionEx, RgFunctionEx
    UpdateCell intISFunction2, cbFunction2
End Sub

Private Sub cbFunctionEx_Change()
    If EnEvents = False Then Exit Sub
        
    Set RgFunctionEx = RangeOfValue(cbFunctionEx.Value, intFunctionEx, RgFunction2)
    UpdateComboBox cbFunctionEx2, RgFunctionEx, intFunctionEx2, RgFunctionEx2
    UpdateCell intISFunctionEx, cbFunctionEx
End Sub

Private Sub cbFunctionEx2_Change()
    If EnEvents = False Then Exit Sub
        
    Set RgFunctionEx2 = RangeOfValue(cbFunctionEx2.Value, intFunctionEx2, RgFunctionEx)
    UpdateComboBox cbFunctionEx3, RgFunctionEx2, intFunctionEx3, RgFunctionEx3
    UpdateCell intISFunctionEx2, cbFunctionEx2
End Sub

Private Sub cbFunctionEx3_Change()
    If EnEvents = False Then Exit Sub
        
    Set RgFunctionEx3 = RangeOfValue(cbFunctionEx3.Value, intFunctionEx3, RgFunctionEx2)
    UpdateCell intISFunctionEx3, cbFunctionEx3
End Sub


'Private Sub cbTopic_Change()
'    If EnEvents = False Then
'        Exit Sub
'    End If
'
'    'If cbTopic.ListIndex = -1 Then Exit Sub
'    UpdateCell intISTopic, cbTopic
'End Sub

Private Sub tbDocuSet_AfterUpdate()
    InputSheet.Cells(intARw, intISDocuSet).Value = tbDocuSet.Value
End Sub

Private Sub UserForm_Activate()
    InputSheet.Activate
End Sub

Private Sub UserForm_Initialize()
    'populates SC items
    dbFormHeight = Me.Height
    Dim cll As Range
    cbSC.Clear
    For Each cll In ws.Range(ws.Cells(intStartRow, intSC), ws.Cells(ws.UsedRange.Rows.Count, intSC))
        If clean(cll.Value) <> "" Then
            AddToCombobox cbSC, cll.Value
        End If
    Next cll
    
    'populate Action and Process Subfolder items
    With cbAction
        .Clear
        .AddItem "Active"
        .AddItem "LEG"
        .AddItem "LEG Secure"
    End With
    'With cbPS
    '    .Clear
    '    .AddItem "Y"
    '    .AddItem "N"
    'End With
    
    Set TBCollection = New Collection
End Sub
