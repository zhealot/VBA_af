VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim rg As Range
    If blWork Then
    
        If Target.Row <> 1 Then
        
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            EnEvents = False
            intARw = Target.Row
            intAClm = Target.Column
    
            Set RgSC = Nothing
            Set RgSiteBranch = Nothing
            Set RgSitePro = Nothing
            Set RgSiteSC = Nothing
            Set RgSite1 = Nothing
            Set RgSite2 = Nothing
            Set RgLibrary = Nothing
            Set RgMeta = Nothing
            UserForm1.cbAction.Value = ""
            GreyCB UserForm1.cbSC
            GreyCB UserForm1.cbSiteBranch
            GreyCB UserForm1.cbSitePro
            GreyCB UserForm1.cbSiteSC
            GreyCB UserForm1.cbSite1
            GreyCB UserForm1.cbSite2
            GreyCB UserForm1.cbSite3
            GreyCB UserForm1.cbLibrary
            GreyCB UserForm1.cbTopic
            With UserForm1.tbDocuSet
                .Value = ""
                .Enabled = False
                .BackColor = &H8000000F
            End With
            Do While TBCollection.Count > 0
                UserForm1.Controls.Remove TBCollection.Item(1).Name
                TBCollection.Remove 1
            Loop
            
            UserForm1.lbFolderID.Caption = clean(InputSheet.Cells(intARw, intISFID)) 'ReadValue(intISFID)
            'UserForm1.lblFolderPath.Caption = clean(InputSheet.Cells(intARw, intISFP)) 'ReadValue(intISFP)
            UserForm1.tbFolderPath.Text = clean(InputSheet.Cells(intARw, intISFP))
            UserForm1.cbAction.SetFocus
            
            'On Error Resume Next
            If ReadValue(intISAction) = "" Then
                GoTo LAST
            'no short code convert for Action
            ElseIf VinCombo(clean(InputSheet.Cells(intARw, intISAction).Value), UserForm1.cbAction) Then   'ReadValue(intISAction), UserForm1.cbAction) Then
                UserForm1.cbAction.Value = clean(InputSheet.Cells(intARw, intISAction).Value)   'ReadValue(intISAction)
                If UserForm1.cbAction.Value <> "Active" Then
                    GoTo LAST
                Else
                    UserForm1.cbSC.Enabled = True
                    UpdateComboBox UserForm1.cbSC, ws.UsedRange, intSC, RgSC
                End If
            Else
                InvalidCB UserForm1.cbAction
                GoTo LAST
            End If
            
            'update SC
            If VinCombo(ReadValue(intISSC), UserForm1.cbSC) Then
                UserForm1.cbSC.Value = ReadValue(intISSC)
                Set RgSC = RangeOfValue(UserForm1.cbSC.Value, intSC, ws.UsedRange)
                UpdateComboBox UserForm1.cbSiteBranch, RgSC, intSiteBranch, RgSiteBranch
            Else
                InvalidCB UserForm1.cbSC
                GoTo LAST
            End If
            
            'update SiteBranch
            If VinCombo(ReadValue(intISSiteBranch), UserForm1.cbSiteBranch) Then
                UserForm1.cbSiteBranch.Value = ReadValue(intISSiteBranch)
                Set RgSiteBranch = RangeOfValue(UserForm1.cbSiteBranch.Value, intSiteBranch, RgSC)
                UserForm1.cbSitePro.Enabled = True
                UpdateComboBox UserForm1.cbSitePro, RgSiteBranch, intSitePro, RgSitePro
                UserForm1.cbSiteSC.Enabled = True
                UpdateComboBox UserForm1.cbSiteSC, RgSiteBranch, intSiteSC, RgSiteSC
                If UserForm1.cbSiteSC.ListCount = 0 Then
                    UpdateComboBox UserForm1.cbSite1, RgSiteBranch, intSite1, RgSite1
                End If
            Else
                InvalidCB UserForm1.cbSiteBranch
                GoTo LAST
            End If
            
            'update SitePro
            If VinCombo(ReadValue(intISSitePro), UserForm1.cbSitePro) Then
                UserForm1.cbSitePro.Value = ReadValue(intISSitePro)
                Set RgSitePro = RangeOfValue(UserForm1.cbSitePro.Value, intSitePro, RgSiteBranch)
            Else
                InvalidCB UserForm1.cbSitePro
                GoTo LAST
            End If
            
            'update SiteSC
            If VinCombo(ReadValue(intISSiteSC), UserForm1.cbSiteSC) Then
                UserForm1.cbSiteSC.Value = ReadValue(intISSiteSC)
                Set RgSiteSC = RangeOfValue(UserForm1.cbSiteSC.Value, intSiteSC, RgSiteBranch)
                UserForm1.cbSite1.Enabled = True
                UserForm1.cbSite2.Enabled = True
                UserForm1.cbSite3.Enabled = True
                UpdateComboBox UserForm1.cbSite1, RgSiteSC, intSite1, RgSite1
                UpdateComboBox UserForm1.cbSite2, RgSiteSC, intSite2, RgSite2
                UpdateComboBox UserForm1.cbSite3, RgSiteSC, intSite3, RgSite3
                UserForm1.cbLibrary.Enabled = True
                If UserForm1.cbSite1.ListCount = 0 Then
                    UpdateComboBox UserForm1.cbLibrary, RgSiteSC, intLibrary, RgLibrary
                Else
                    If RgSiteSC Is RgSite1 Then
                        UpdateComboBox UserForm1.cbLibrary, RgSiteSC, intLibrary, RgLibrary
                    ElseIf Intersect(MinusRg(RgSiteSC, RgSite1), ws.Columns(intLibrary)) Is Nothing Then
                        UpdateComboBox UserForm1.cbLibrary, RgSiteSC, intLibrary, RgLibrary
                    Else
                        UpdateComboBox UserForm1.cbLibrary, MinusRg(RgSiteSC, RgSite1), intLibrary, RgLibrary
                    End If
                End If
            Else
                InvalidCB UserForm1.cbSiteSC
                GoTo LAST
            End If
            
            'update Site1
            If VinCombo(ReadValue(intISSite1), UserForm1.cbSite1) Then
                UserForm1.cbSite1.Value = ReadValue(intISSite1)
                Set RgSite1 = RangeOfValue(UserForm1.cbSite1.Value, intSite1, RgSiteSC)
                UpdateComboBox UserForm1.cbSite2, RgSite1, intSite2, RgSite2
                UpdateComboBox UserForm1.cbSite3, RgSite2, intSite3, RgSite3
                UpdateComboBox UserForm1.cbLibrary, RgSite1, intLibrary, RgLibrary
            Else
                InvalidCB UserForm1.cbSite1
                GoTo LAST
            End If
            
            'update Site2
            If VinCombo(ReadValue(intISSite2), UserForm1.cbSite2) Then
                UserForm1.cbSite2.Value = ReadValue(intISSite2)
                Set RgSite2 = RangeOfValue(UserForm1.cbSite2.Value, intSite2, RgSite1)
                UpdateComboBox UserForm1.cbLibrary, RgSite2, intLibrary, RgLibrary
            Else
                InvalidCB UserForm1.cbSite2
                GoTo LAST
            End If
                        
            'update Site3
            If VinCombo(ReadValue(intISSite3), UserForm1.cbSite3) Then
                UserForm1.cbSite3.Value = ReadValue(intISSite3)
                Set RgSite3 = RangeOfValue(UserForm1.cbSite3.Value, intSite3, RgSite2)
                UpdateComboBox UserForm1.cbLibrary, RgSite3, intLibrary, RgLibrary
            Else
                InvalidCB UserForm1.cbSite3
                GoTo LAST
            End If
            
            'update Library
            If VinCombo(ReadValue(intISLibrary), UserForm1.cbLibrary) Then
                UserForm1.cbLibrary.Value = ReadValue(intISLibrary)
                Set RgLibrary = RangeOfValue(UserForm1.cbLibrary.Value, intLibrary, RgSite2)
                UserForm1.cbTopic.Enabled = True
                UpdateComboBox UserForm1.cbTopic, RgLibrary, intTopic, RgTopic
                
                'update Meta data
                UpdateComboBox UserForm1.cbTemp, RgLibrary, intMeta, RgMeta
                Dim str As String
                Dim tbMetas As MSForms.TextBox
                Dim i As Long
                Dim sA() As String
                Dim tmpA() As String
                UserForm1.Height = dbFormHeight
                If UserForm1.cbTemp.ListCount > 0 Then
                    For i = 0 To UserForm1.cbTemp.ListCount - 1
                        Set tbMetas = AddControl("Forms.TextBox.1", "tbMetaKey" & i, UserForm1.lbMeta.Top + i * 16, UserForm1.tbDocuSet.Left, 16, 100)
                        TBCollection.Add tbMetas, tbMetas.Name
                        tbMetas.Value = UserForm1.cbTemp.List(i)
                        tbMetas.Enabled = False
                        tbMetas.Font.Size = 8
                        Set tbMetas = AddControl("Forms.TextBox.1", "tbMetaValue" & i, UserForm1.lbMeta.Top + i * 16, UserForm1.tbDocuSet.Left + 110, 16, 100)
                        tbMetas.Font.Size = 8
                        TBCollection.Add tbMetas, tbMetas.Name
                        UserForm1.Height = UserForm1.Height + 16
                    Next i
                End If
                'if exist meta data
                str = ReadValue(intISMeta)
                If Right(Trim(str), 1) = "," Then
                    str = Left(Trim(str), Len(Trim(str)) - 1)
                End If
                If str <> "" Then
                    UserForm1.Height = dbFormHeight
                    Do While TBCollection.Count > 0
                        UserForm1.Controls.Remove TBCollection.Item(1).Name
                        TBCollection.Remove 1
                    Loop
                    sA = Split(str, ",")
                    For i = 0 To UBound(sA)
                        tmpA = Split(sA(i), ":")
                        Set tbMetas = AddControl("Forms.TextBox.1", "tbMetaKey" & i, UserForm1.lbMeta.Top + i * 16, UserForm1.tbDocuSet.Left, 16, 100)
                        TBCollection.Add tbMetas, tbMetas.Name
                        tbMetas.Enabled = False
                        tbMetas.Value = Trim(tmpA(0))
                        tbMetas.Font.Size = 8
                        Set tbMetas = AddControl("Forms.TextBox.1", "tbMetaValue" & i, UserForm1.lbMeta.Top + i * 16, UserForm1.tbDocuSet.Left + 110, 16, 100)
                        TBCollection.Add tbMetas, tbMetas.Name
                        tbMetas.Value = Trim(IIf(UBound(tmpA) > 0, tmpA(1), ""))
                        tbMetas.Font.Size = 8
                        UserForm1.Height = UserForm1.Height + 16
                    Next i
                End If 'If str <> "" Then
                
                'update Docu Set
                Set rg = Intersect(ws.Range(ws.Cells(intStartRow, intLibrary), ws.Cells(ws.UsedRange.Rows.Count, intLibrary)), RgLibrary).Find(UserForm1.cbLibrary.Value, LookIn:=xlValues, lookat:=xlWhole)
                If Not rg Is Nothing Then
                    If rg.Offset(0, 1).Value = "Yes" Then
                        With UserForm1.tbDocuSet
                            .Enabled = True
                            .BackColor = &H80000005
                            .Value = ReadValue(intISDocuSet)
                        End With
                    End If
                End If
            Else
               InvalidCB UserForm1.cbLibrary
                GoTo LAST
            End If
            
            'update Topic
            If VinCombo(ReadValue(intISTopic), UserForm1.cbTopic) Then
                UserForm1.cbTopic.Value = ReadValue(intISTopic)
            Else
                InvalidCB UserForm1.cbTopic
                GoTo LAST
            End If
        Else    'if target.row =1
            On Error Resume Next
            If Target.Columns.Count = 1 And Target.Rows.Count = 1 Then
                Target.Offset(1, 0).Activate
            End If
            On Error GoTo 0
        End If
LAST:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    EnEvents = True
    
    End If 'If blWork And Target.Row <> 1
End Sub
