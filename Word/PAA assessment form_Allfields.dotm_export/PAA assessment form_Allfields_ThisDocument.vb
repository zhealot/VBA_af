VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    Dim doc As Document
    Dim cc As ContentControl 'source ContentControl
    Dim ccTmp As ContentControl 'target ContentControl
        
    Set doc = ActiveDocument
    Set cc = ContentControl
    
    'disable paragraph marks
    doc.ActiveWindow.ActivePane.View.ShowAll = False
    
    If cc.Title = "Committee" Then
        If doc.SelectContentControlsByTitle("fromCommittee").Count > 0 Then
            EditCC doc.SelectContentControlsByTitle("fromCommittee").Item(1), cc.Range.Text
        End If
    End If

    If cc.Title = "Team" Then
        Dim ety
        If cc.Range.Text = "P&I Investment Advisor" Then
            cc.LockContents = False
            cc.LockContentControl = False
            Set cc = doc.SelectContentControlsByTitle("SupportedBy").Item(1)
            cc.SetPlaceholderText , , " "
            For Each ety In cc.DropdownListEntries
                If ety.Value = "" Then
                    ety.Select
                    Exit For
                End If
            Next ety
            Set cc = doc.SelectContentControlsByTitle("AssessedBy").Item(1)
            cc.SetPlaceholderText , , " "
            For Each ety In cc.DropdownListEntries
                If ety.Value = "" Then
                    ety.Select
                    Exit For
                End If
            Next ety
            cc.LockContents = True
            cc.LockContentControl = True
            HideBm "bmPeerReview", "bmPeerReview", True
        End If
        If cc.Range.Text = "Outcome Planning" Then
            Set cc = doc.SelectContentControlsByTitle("SupportedBy").Item(1)
            cc.LockContents = False
            cc.SetPlaceholderText , , "Choose supported by person"
            For Each ety In cc.DropdownListEntries
                If ety.Value = "" Then
                    ety.Select
                    Exit For
                End If
            Next ety
            Set cc = doc.SelectContentControlsByTitle("AssessedBy").Item(1)
            cc.LockContents = False
            cc.SetPlaceholderText , , "Choose assessed by person"
            For Each ety In cc.DropdownListEntries
                If ety.Value = "" Then
                    ety.Select
                    Exit For
                End If
            Next ety
            HideBm "bmPeerReview", "bmPeerReview", False
        End If
    End If
            
    If cc.Title = "Amount" Then
        If ParseNumber(cc.Range.Text) > 0 Then
            cc.Range.Text = Format(ParseNumber(cc.Range.Text), "###,###,##0.00")    'format number to currency
        End If
        If IsNumeric(cc.Range.Text) And doc.SelectContentControlsByTitle("Team").Item(1).Range.Text = "Outcome Planning" Then
            Dim amt As Double
            amt = cc.Range.Text
            If amt >= 15000000 Then
                If doc.SelectContentControlsByTitle("SupportedBy").Count > 0 Then
                    Set cc = doc.SelectContentControlsByTitle("SupportedBy").Item(1)
                    cc.LockContents = False
                    Dim entry
                    For Each entry In cc.DropdownListEntries
                        If entry.Value = "Neil Cree" Then
                            entry.Select
                            Exit For
                        End If
                    Next entry
                    cc.LockContents = True
                End If
            Else
                If doc.SelectContentControlsByTitle("SupportedBy").Count > 0 Then
                    Set cc = doc.SelectContentControlsByTitle("SupportedBy").Item(1)
                    cc.LockContents = False
                    cc.DropdownListEntries(1).Select
                End If
            End If
        End If
    End If
    
    If cc.Title = "Recommendation" Then
        If InStr(cc.Range.Text, "conditions") > 0 Then
            If doc.Bookmarks.Exists("bmRecommendation") Then
                doc.Bookmarks("bmRecommendation").Select
            End If
        End If
    End If
    
    If cc.Title = "AssessedBy" Then
        If InStr(doc.SelectContentControlsByTitle("Recommendation").Item(1).Range.Text, "conditions") > 0 Then
            If Trim(doc.Bookmarks("bmRecommendation").Range.Cells(1).Range.Text) = Chr(13) & Chr(7) Then
                cc.DropdownListEntries.Item(1).Select
                MsgBox "Please enter conditions"
                doc.Bookmarks("bmRecommendation").Select
            End If
        End If
    End If
    
    
    If cc.Title = "Phases" Then
        Select Case cc.Range.Text
            Case "Strategic case"
                HideBm "bmReasonForRecommendation", "bmReasonForRecommendation", True
                HideBm "bmBenefitAndCost", "bmBenefitAndCost", True
                HideBm "bmIncrementalAnalysis", "bmIncrementalAnalysis", True
                HideBm "bmSensitivityTesting", "bmSensitivityTesting", True
                HideBm "bmPhaseCashStart", "bmConstructionCost", True
                HideBm "bmConstructionCost", "bmReasonForAny", True
                HideBm "bmAlternatives", "bmAlternatives", True
                HideBm "bmDetailedProgramme", "bmDetailedProgramme", True
            Case "Programme business case"
                HideBm "bmReasonForRecommendation", "bmReasonForRecommendation", False
                HideBm "bmBenefitAndCost", "bmBenefitAndCost", False
                If doc.Bookmarks.Exists("bmBenefitAndCost") Then
                    If doc.Bookmarks("bmBenefitAndCost").Range.Cells.Count > 0 Then
                        doc.Bookmarks("bmBenefitAndCost").Range.Cells(1).Range.Text = "Indicative benefit and cost appraisal"
                    End If
                End If
                HideBm "bmIncrementalAnalysis", "bmIncrementalAnalysis", True
                HideBm "bmSensitivityTesting", "bmSensitivityTesting", True
                HideBm "bmPhaseCashStart", "bmConstructionCost", False
                HideBm "bmConstructionCost", "bmReasonForAny", True
                HideBm "bmAlternatives", "bmAlternatives", False
                HideBm "bmDetailedProgramme", "bmDetailedProgramme", True
            Case "Indicative business case"
                HideBm "bmReasonForRecommendation", "bmReasonForRecommendation", False
                HideBm "bmBenefitAndCost", "bmBenefitAndCost", False
                If doc.Bookmarks.Exists("bmBenefitAndCost") Then
                    If doc.Bookmarks("bmBenefitAndCost").Range.Cells.Count > 0 Then
                        doc.Bookmarks("bmBenefitAndCost").Range.Cells(1).Range.Text = "Benefit and cost appraisal"
                    End If
                End If
                HideBm "bmIncrementalAnalysis", "bmIncrementalAnalysis", False
                HideBm "bmSensitivityTesting", "bmSensitivityTesting", False
                HideBm "bmPhaseCashStart", "bmConstructionCost", False
                HideBm "bmConstructionCost", "bmReasonForAny", False
                HideBm "bmAlternatives", "bmAlternatives", False
                HideBm "bmDetailedProgramme", "bmDetailedProgramme", False
            Case "Detailed business case"
                HideBm "bmReasonForRecommendation", "bmReasonForRecommendation", False
                HideBm "bmBenefitAndCost", "bmBenefitAndCost", False
                If doc.Bookmarks.Exists("bmBenefitAndCost") Then
                    If doc.Bookmarks("bmBenefitAndCost").Range.Cells.Count > 0 Then
                        doc.Bookmarks("bmBenefitAndCost").Range.Cells(1).Range.Text = "Benefit and cost appraisal"
                    End If
                End If
                HideBm "bmIncrementalAnalysis", "bmIncrementalAnalysis", False
                HideBm "bmSensitivityTesting", "bmSensitivityTesting", False
                HideBm "bmPhaseCashStart", "bmConstructionCost", False
                HideBm "bmConstructionCost", "bmReasonForAny", False
                HideBm "bmAlternatives", "bmAlternatives", False
                HideBm "bmDetailedProgramme", "bmDetailedProgramme", False
'            Case "Pre-implementation"
'                HideBm "bmReasonForRecommendation", "bmReasonForRecommendation", False
'                HideBm "bmBenefitAndCost", "bmBenefitAndCost", False
'                HideBm "bmIncrementalAnalysis", "bmIncrementalAnalysis", False
'                HideBm "bmSensitivityTesting", "bmSensitivityTesting", False
'                HideBm "bmPhaseCashStart", "bmConstructionCost", False
'                HideBm "bmConstructionCost", "bmReasonForAny", True
'                HideBm "bmAlternatives", "bmAlternatives", False
'                HideBm "bmDetailedProgramme", "bmDetailedProgramme", False
'            Case "Implementation"
'                HideBm "bmReasonForRecommendation", "bmReasonForRecommendation", False
'                HideBm "bmBenefitAndCost", "bmBenefitAndCost", False
'                HideBm "bmIncrementalAnalysis", "bmIncrementalAnalysis", False
'                HideBm "bmSensitivityTesting", "bmSensitivityTesting", False
'                HideBm "bmPhaseCashStart", "bmConstructionCost", False
'                HideBm "bmConstructionCost", "bmReasonForAny", True
'                HideBm "bmAlternatives", "bmAlternatives", False
'                HideBm "bmDetailedProgramme", "bmDetailedProgramme", False
            Case Else
                HideBm "bmReasonForRecommendation", "bmReasonForRecommendation", False
                HideBm "bmBenefitAndCost", "bmBenefitAndCost", False
                HideBm "bmIncrementalAnalysis", "bmIncrementalAnalysis", False
                HideBm "bmSensitivityTesting", "bmSensitivityTesting", False
                HideBm "bmPhaseCashStart", "bmConstructionCost", False
                HideBm "bmConstructionCost", "bmReasonForAny", False
                HideBm "bmAlternatives", "bmAlternatives", False
                HideBm "bmDetailedProgramme", "bmDetailedProgramme", False
        End Select
    End If  'if cc.Title="Phases"
End Sub

Function EditCC(ByRef cc As ContentControl, str As String)
    Dim blLockEdit As Boolean
    blLockEdit = cc.LockContents
    cc.LockContents = False
    cc.Range.Text = str
    cc.LockContents = blLockEdit
End Function

Function ParseNumber(str As String) As Double
    ParseNumber = -1
    str = Trim(str)
    
    If str = "" Then Exit Function
    If IsNumeric(str) Then
        ParseNumber = str
        Exit Function
    End If
    
    str = Replace(str, "$", "")
    If EndWith(str, "m") Or EndWith(str, "mil") Or EndWith(str, "million") Then
        If EndWith(str, "m") Then
            str = Left(str, Len(str) - 1)
        ElseIf EndWith(str, "mil") Then
            str = Left(str, Len(str) - 3)
        ElseIf EndWith(str, "million") Then
            str = Left(str, Len(str) - 7)
        End If
        If IsNumeric(str) Then
            If str > 0 Then
                ParseNumber = str * 1000000
            End If
        End If
    ElseIf EndWith(str, "b") Or EndWith(str, "bil") Or EndWith(str, "billion") Then
        If EndWith(str, "b") Then
            str = Left(str, Len(str) - 1)
        ElseIf EndWith(str, "bil") Then
            str = Left(str, Len(str) - 3)
        ElseIf EndWith(str, "billion") Then
            str = Left(str, Len(str) - 7)
        End If
        If IsNumeric(str) Then
            If str > 0 Then
                ParseNumber = str * 1000000000
            End If
        End If
    End If
End Function

Function EndWith(sIn As String, sEnd As String) As Boolean
    EndWith = False
    If Len(sEnd) >= Len(sIn) Then Exit Function
    If LCase(Right(sIn, Len(sEnd))) = LCase(sEnd) Then
        EndWith = True
    End If
End Function

Sub test()
    Debug.Print Format(ParseNumber("2.4dfwe"), "###,###,##0.00")
End Sub

Function HideBm(F As String, T As String, blHide As Boolean)
'hide rows in a table
    Dim doc As Document
    Set doc = ActiveDocument
    
    If Not doc.Bookmarks.Exists(F) Or Not doc.Bookmarks.Exists(T) Then Exit Function
    Dim bmFrom As Bookmark
    Dim bmTo As Bookmark
    Set bmFrom = doc.Bookmarks(F)
    Set bmTo = doc.Bookmarks(T)
    If bmFrom.Range.Tables.Count = 0 Or bmTo.Range.Tables.Count = 0 Then Exit Function
    
    On Error GoTo E5592
    Dim iRw As Integer
    iRw = bmFrom.Range.Rows(1).Index
    'same start/end bookmark, hide one row
    If bmFrom = bmTo And iRw <> 1 Then
        bmFrom.Range.Tables(1).Rows(bmFrom.Range.Rows(1).Index).Range.Font.Hidden = blHide
    'start bookmark in row 1, hide whole table
    ElseIf iRw = 1 Then
        bmFrom.Range.Tables(1).Range.Font.Hidden = blHide
    Else
E5592:  If Not bmFrom = bmTo Then
            Dim rg As Range
            Set rg = doc.Range
            rg.SetRange bmFrom.Range.Start - 1, bmTo.Range.Start
            rg.Font.Hidden = blHide
        End If
    End If
End Function

Sub setCC()
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If cc.Title = "Charges" Then
            cc.SetPlaceholderText , , "Please choose"
        End If
    Next cc
End Sub
