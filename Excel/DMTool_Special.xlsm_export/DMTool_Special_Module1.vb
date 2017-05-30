Attribute VB_Name = "Module1"
Option Explicit

'information sheet columns
Public Const intSC = 1
Public Const intSiteBranch = 2
Public Const intSitePro = 3
Public Const intSiteSC = 4
Public Const intSite1 = 5
Public Const intSite2 = 6
Public Const intSite3 = 7
Public Const intLibrary = 8
Public Const intDocuSet = 9
Public Const intFunction = 10
Public Const intFunction2 = 11
Public Const intFunctionEx = 12
Public Const intFunctionEx2 = 13
Public Const intFunctionEx3 = 14
Public Const intMeta = 15
Public Const intStartRow = 2
'Input Sheet columns
Public intARw As Long
Public intAClm As Long
Public Const intISFID = 1
Public Const intISFP = 2
Public Const intISPS = 3
Public Const intISAction = 4
Public Const intISSC = 5
Public Const intISSiteBranch = 6
Public Const intISSitePro = 7
Public Const intISSiteSC = 8
Public Const intISSite1 = 9
Public Const intISSite2 = 10
Public Const intISSite3 = 11
Public Const intISLibrary = 12
Public Const intISDocuSet = 13
Public Const intISFunction = 14
Public Const intISFunction2 = 15
Public Const intISFunctionEx = 16
Public Const intISFunctionEx2 = 17
Public Const intISFunctionEx3 = 18
Public Const intISMeta = 19

Public TBCollection As Collection
Public TBArray() As New control
Public ws As Worksheet
Public InputSheet As Worksheet
Public SCSheet As Worksheet

Public blWork As Boolean
Public dbFormHeight As Double

Public RgSC As Range
Public RgSiteBranch As Range
Public RgSitePro As Range
Public RgSiteSC As Range
Public RgSite1 As Range
Public RgSite2 As Range
Public RgSite3 As Range
Public RgLibrary As Range
Public RgFunction As Range
Public RgFunction2 As Range
Public RgFunctionEx As Range
Public RgFunctionEx2 As Range
Public RgFunctionEx3 As Range
Public RgMeta As Range
Public Const SCClm = 5  'short code column
Public Const NullValue = "/"

Public EnEvents As Boolean


Sub Start()
    'Set ws = Application.Workbooks(WorkbookName).Sheets(1)
    'Set InputSheet = Application.Workbooks(WorkbookName).Sheets(2)
    Set ws = ThisWorkbook.Sheets("Ref")
    Set SCSheet = ThisWorkbook.Sheets("Shortcodes")
    Set InputSheet = ThisWorkbook.Sheets(3)
    If UserForm1.Visible = False Then
        UserForm1.Show
    End If
End Sub

Sub RibbonButton(control As IRibbonControl)
    'Set ws = Application.Workbooks(WorkbookName).Sheets(1)
    'Set InputSheet = Application.Workbooks(WorkbookName).Sheets(2)
    Set ws = ThisWorkbook.Sheets("Ref")
    Set SCSheet = ThisWorkbook.Sheets("Shortcodes")
    Set InputSheet = ThisWorkbook.Sheets(3)
    If UserForm1.Visible = True Then
        blWork = False
        UserForm1.Hide
    Else
        blWork = True
        UserForm1.Show
        InputSheet.Parent.Activate
        InputSheet.Activate
        On Error Resume Next
        ActiveCell.Offset(0, 1).Activate
        'On Error GoTo 0
    End If
End Sub

'read string from active row, no non-print characters and trimmed
Public Function ReadValue(clm As Integer) As String
    Dim rg As Range
    Dim str As String
    Dim i As Integer
    str = clean(InputSheet.Cells(intARw, clm).Value)
    ReadValue = str
    Set rg = Nothing
    If str <> "" Then
        Set rg = SCSheet.Columns(SCClm).Find(str, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)        '.UsedRange.Find(str, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
        If rg Is Nothing Then
            Exit Function
        Else
            If rg.Column = SCClm Then
                For i = SCClm - 1 To 1 Step -1
                    If SCSheet.Cells(rg.Row, i).Value <> "" Then
                        ReadValue = SCSheet.Cells(rg.Row, i).Value
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
End Function

'add an item to combobox if it's not there
Sub AddToCombobox(cb As MSForms.ComboBox, str As String)
    If str = "" Then Exit Sub
    
    Dim i As Integer
    Dim found As Boolean
    found = False
    For i = 0 To cb.ListCount - 1
        If cb.List(i) = str Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then cb.AddItem str
End Sub

Sub UpdateComboBox(cb As MSForms.ComboBox, rg As Range, clm As Long, vRg As Range)
' update combobox and range of the combobox
    
    Dim cll As Range
    Dim tmpRg As Range
    Dim i As Long
    
    Set vRg = Nothing
    
    'Debug.Print ("CB name: " & cb.Name)
    Application.EnableEvents = False
    cb.Clear
    Set tmpRg = Intersect(rg, ws.Range(ws.Cells(intStartRow, clm), ws.Cells(ws.UsedRange.Rows.Count, clm)))
    If Not tmpRg Is Nothing Then
        If clean(tmpRg.Cells(1, 1).Value) = "" Then
            AddToCombobox cb, NullValue
        End If
        For Each cll In tmpRg
            If clean(cll.Value <> "") Then
                AddToCombobox cb, clean(cll.Value)
            End If
        Next cll
    End If
    With cb
        If (.ListCount = 0) Or (.ListCount = 1 And .List(0) = NullValue) Then
            .Clear
            .Enabled = False
        Else
            .Enabled = True
        End If
    End With
    Set vRg = RangeOfValue("", clm, rg)
    Application.EnableEvents = True
    
End Sub

Sub UpdateComboBox_Old(cb As MSForms.ComboBox, rg As Range, clm As Long, vRg As Range)
' update combobox and range of the combobox
    Dim cll As Range
    Dim tmpRg As Range
    Dim i As Long
    
    Set vRg = Nothing
    
    Application.EnableEvents = False
    cb.Clear
    Set tmpRg = Intersect(rg, ws.Range(ws.Cells(intStartRow, clm), ws.Cells(ws.UsedRange.Rows.Count, clm)))
    If Not tmpRg Is Nothing Then
        
        For Each cll In tmpRg
            If clean(cll.Value <> "") Then
                AddToCombobox cb, clean(cll.Value)
            End If
        Next cll
    End If
    If (cb.ListCount = 0) Then
        Set vRg = RangeOfValue("", clm, rg)
        cb.Enabled = False
    Else
        cb.Enabled = True
        For i = 0 To cb.ListCount - 1
            If vRg Is Nothing Then
                Set vRg = RangeOfValue(cb.List(i), clm, rg)
            Else
                Set vRg = Union(vRg, RangeOfValue(cb.List(i), clm, rg))
            End If
        Next i
    End If
    Application.EnableEvents = True
End Sub


Sub UpdateCell(clm As Integer, cb As MSForms.ComboBox)
    InputSheet.Activate
    If cb.Value = NullValue Then
        InputSheet.Cells(intARw, clm).Value = NullValue '###
    ElseIf cb.Value = " " Then
        InputSheet.Cells(intARw, clm).Value = ""
    Else
        InputSheet.Cells(intARw, clm).Value = IIf(ShortCode(cb.Value) = "", cb.Value, ShortCode(cb.Value))
    End If
    Call ClearTail(clm)
End Sub

Public Function UpdateMeta() As Integer
    Dim i As Integer
    Dim str As String
    If TBCollection.Count <= 0 Then
        UpdateMeta = 1
        Exit Function
    End If
    'check for "," and ":"
    For i = 1 To TBCollection.Count
        If InStr(TBCollection.Item(i).Value, ",") > 0 Or InStr(TBCollection.Item(i).Value, ":") > 0 Then
            MsgBox "Please do not use colon or comma in meta data"
            UpdateMeta = 0
            Exit Function
        End If
    Next i
    'get string, if no value then no need to write the key for that value
    str = ""
    i = 1
    Do While i < TBCollection.Count
        If Trim(clean(TBCollection.Item(i + 1).Value)) <> "" Then
            str = str & TBCollection.Item(i).Value & ":" & TBCollection.Item(i + 1).Value & ", "
        End If
        i = i + 2
    Loop
    If Right(str, 2) = ", " Then
        str = Left(str, Len(str) - 2)
    End If
    InputSheet.Cells(intARw, intISMeta).Value = str
    UpdateMeta = 1
End Function

Public Function AddControl(CType As String, nm As String, t As Double, l As Double, h As Double, w As Double) As control
        Dim ctrl As control
        Set ctrl = UserForm1.Controls.Add(CType, nm)
        With ctrl
            .Top = t
            .Left = l
            .Height = h
            .Width = w
        End With
        Set AddControl = ctrl
End Function

Public Function clean(str As String) As String
    clean = Trim(WorksheetFunction.clean(str))
End Function

Public Function RangeOfValue(v As String, c As Long, rg As Range) As Range
'set range based on value and lookin-range
    Set RangeOfValue = Nothing
    v = clean(v)
    If v = "" Then
        Set RangeOfValue = rg
        Exit Function
    End If
    
    Dim LookInRg As Range
    Dim tmp As Range
    Dim FirstFound As Long
    Set tmp = Intersect(ws.Range(ws.Cells(intStartRow, c), ws.Cells(ws.UsedRange.Rows.Count, c)), rg)
    If v = NullValue Then
        Set RangeOfValue = ws.Range(tmp.Cells(1, 1), ws.Cells(tmp.Cells(1, 1).End(xlDown).Row - 1, ws.UsedRange.Columns.Count))
    Else
        Set LookInRg = tmp.Find(v, LookIn:=xlValues, lookat:=xlWhole)
        If Not LookInRg Is Nothing Then
            FirstFound = LookInRg.Row
            Set RangeOfValue = LookInRg
            Do
                If LookInRg.Offset(1, 0).Value <> "" Then
                    Set RangeOfValue = Union(RangeOfValue, ws.Range(LookInRg, ws.Cells(LookInRg.Row, ws.UsedRange.Columns.Count)))
                Else
                    Set RangeOfValue = Union(RangeOfValue, ws.Range(LookInRg, ws.Cells(WorksheetFunction.Min(LookInRg.End(xlDown).Row - 1, tmp.Areas(tmp.Areas.Count).Row + tmp.Areas(tmp.Areas.Count).Rows.Count - 1), ws.UsedRange.Columns.Count)))
                End If
                Set LookInRg = tmp.FindNext(LookInRg)
            Loop While LookInRg.Row <> FirstFound
        End If
    End If
    
    
End Function

Sub testrg()
    Dim rr As Range
    Dim frm As Range
    Set ws = ThisWorkbook.Sheets(2)
    Set frm = ws.Range(ws.Cells(211, 5), ws.Cells(239, 10))
    Set rr = RangeOfValue("Annual Report", 7, frm)
    rr.Select
End Sub

Public Function RgInRg(rg1 As Range, rg2 As Range) As Boolean
    If Intersect(rg1, rg2) Is Nothing Then
        RgInRg = False
    Else
        RgInRg = True
    End If
End Function

Public Function MinusRg(fromRg As Range, rg As Range) As Range
    Dim cl As Range
    Set MinusRg = Nothing
    If Intersect(fromRg, rg) Is Nothing Then
        Set MinusRg = fromRg
        Exit Function
    End If
    If fromRg Is rg Then
        Set MinusRg = Nothing
        Exit Function
    End If
    
    For Each cl In fromRg
        If Intersect(cl, rg) Is Nothing Then
            If MinusRg Is Nothing Then
                Set MinusRg = cl
            Else
                Set MinusRg = Union(MinusRg, cl)
            End If
        End If
    Next cl
End Function



Public Function GreyCB(cb As MSForms.ComboBox)
    cb.Clear
    cb.Enabled = False
    cb.Value = ""
End Function

Public Function ShortCode(str As String) As String
    ShortCode = ""
    str = clean(str)
    If str = "" Then Exit Function
         
    Dim rg As Range
    Set rg = SCSheet.UsedRange.Find(str, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)    '.Columns(SCClm).Find(str, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
    If Not rg Is Nothing Then
        If rg.Column <> SCClm Then
            ShortCode = SCSheet.Cells(rg.Row, SCClm).Value
        End If
    End If
End Function

Public Function VinCombo(str As String, cb As MSForms.ComboBox) As Boolean
    VinCombo = False
    If cb.ListCount <= 0 Then
        If str = "" Then
            VinCombo = True
            Exit Function
        End If
    Else
        Dim i As Integer
        For i = 0 To cb.ListCount - 1
        If str = cb.List(i) Then
            VinCombo = True
            Exit Function
        End If
        Next i
    End If
End Function

'### validation not needed by users
Public Function InvalidCB(cb As MSForms.ComboBox)
    'cb.Value = ""
    'MsgBox "Invalid value in: " & Replace(cb.Name, "cb", "")
    'cb.SetFocus
End Function

Public Function DisCB(cb As MSForms.ComboBox, clm As Integer)
    cb.Value = ""
    cb.Enabled = False
    UpdateCell clm, cb
End Function

Sub TidyUpRef()
    Application.ScreenUpdating = False
    Dim cl As Range
    Dim st As Worksheet
    Set st = ActiveWorkbook.Sheets("Ref")
    For Each cl In st.UsedRange.Cells
       cl.Value = clean(cl.Value)
    Next cl
    Application.ScreenUpdating = True
End Sub


Sub ShortcodeInRef()
    Dim sSC As Worksheet
    Dim sRef As Worksheet
    Dim WB As Workbook
    Set WB = ThisWorkbook
    Set sSC = WB.Sheets("Shortcodes")
    Set sRef = WB.Sheets("Ref")
    
    Dim cl As Range
    Dim rg As Range
    For Each cl In sSC.UsedRange.Cells
        If clean(cl.Value) <> "" Then
            Set rg = sRef.UsedRange.Find(clean(cl.Value), LookIn:=xlValues, lookat:=xlWhole)
            If rg Is Nothing Then
                cl.Interior.ColorIndex = 4
            End If
        End If
    Next
End Sub

Function ClearTail(clm As Integer)
    If intAClm < intISMeta Then
        InputSheet.Range(Cells(intARw, clm + 1), Cells(intARw, intISMeta)).Clear
    End If
End Function
