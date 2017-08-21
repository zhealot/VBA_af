VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Built by Allfields for Harrison Grierson, to be used on Windows with Microsoft Office 2010/2013/2016
'Last edited:   22/08/2017, tao@allfields.co.nz
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    Application.ScreenUpdating = False
    On Error Resume Next
    If ContentControl.Tag = "ccMonth" Then    'for cc Month, caculate correct calendar layout
        If Not IsDate(ContentControl.Range.Text) Then Exit Sub
        On Error Resume Next
        Dim cc As ContentControl
        Dim doc As Document
        Dim sDate As String
        Dim tb As Table
        
        sDate = ContentControl.Range.Text
        Set doc = ThisDocument
        'validate input data
        Dim iDays As Integer    'number of days of the selected month
        Dim iLastMonthDays As Integer   'number of days of last month
        iDays = DaysOfMonth(CDate(sDate))
        iLastMonthDays = DaysOfMonth(DateSerial(Year(IIf(Month(sDate) = "1", Year(sDate) - 1, Year(sDate))), IIf(Month(sDate) = "1", "12", Month(sDate) - 1), 1))
        If Err.Number > 0 Then Exit Sub     'non-date input
                
        'set Month and Year text
        ContentControl.Range.Text = Format(sDate, "MMMM")
        If doc.SelectContentControlsByTag("ccYear").Count > 0 Then
            Set cc = doc.SelectContentControlsByTag("ccYear").Item(1)
            cc.Range.Text = Format(sDate, "yyyy")
        End If
        If doc.SelectContentControlsByTag("ccDate").Count > 0 Then
            doc.SelectContentControlsByTag("ccDate").Item(1).Range.Text = Format(sDate, "MMMM yyyy")
        End If
        
        ' fill calendar's day numbers, based on month chosen
        For Each tb In doc.Tables
            If InStr(LCase(tb.Cell(4, 1).Range.Text), "mon") > 0 Then
                If tb.Rows.Count = 16 And tb.Columns.Count = 7 Then
                    Dim iRw As Integer
                    Dim cl As Cell
                    Dim iClm As Integer
                    Dim iCounter As Integer
                    Dim iFirstWeekDay As Integer    'what day is the first day of the month
                                        
                    iCounter = 1
                    iFirstWeekDay = Format(DateSerial(Year(sDate), Month(sDate), 1), "w")
                    iFirstWeekDay = IIf(iFirstWeekDay = "1", "7", iFirstWeekDay - 1)    'make Monday=1, Sunday=7
                    For iRw = 5 To 15 Step 2
                        For iClm = 1 To 7 Step 1
                            If iRw = 5 Then
                                If iClm >= iFirstWeekDay And iCounter <= iDays Then
                                    tb.Cell(iRw, iClm).Range.Text = iCounter
                                    tb.Cell(iRw, iClm).Range.Font.Color = wdColorBlack
                                    tb.Cell(iRw, iClm).Shading.ForegroundPatternColor = wdColorWhite
                                    tb.Cell(iRw + 1, iClm).Range.Font.Color = wdColorBlack
                                    tb.Cell(iRw + 1, iClm).Shading.ForegroundPatternColor = wdColorWhite
                                    iCounter = iCounter + 1
                                Else
                                    'fill cell with last month's last days
                                    tb.Cell(iRw, iClm).Range.Text = iLastMonthDays + iClm - iFirstWeekDay + 1
                                    tb.Cell(iRw, iClm).Range.Font.Color = wdColorGray40
                                    tb.Cell(iRw, iClm).Shading.ForegroundPatternColor = wdColorGray10
                                    tb.Cell(iRw + 1, iClm).Range.Font.Color = wdColorGray40
                                    tb.Cell(iRw + 1, iClm).Shading.ForegroundPatternColor = wdColorGray10
                                End If
                            Else
                                If iCounter <= iDays Then
                                    tb.Cell(iRw, iClm).Range.Text = iCounter
                                    tb.Cell(iRw, iClm).Range.Font.Color = wdColorBlack
                                    tb.Cell(iRw, iClm).Shading.ForegroundPatternColor = wdColorWhite
                                    tb.Cell(iRw + 1, iClm).Range.Font.Color = wdColorBlack
                                    tb.Cell(iRw + 1, iClm).Shading.ForegroundPatternColor = wdColorWhite
                                Else
                                    tb.Cell(iRw, iClm).Range.Text = iCounter - iDays
                                    tb.Cell(iRw, iClm).Range.Font.Color = wdColorGray40
                                    tb.Cell(iRw, iClm).Shading.ForegroundPatternColor = wdColorGray10
                                    tb.Cell(iRw + 1, iClm).Range.Font.Color = wdColorGray40
                                    tb.Cell(iRw + 1, iClm).Shading.ForegroundPatternColor = wdColorGray10
                                End If
                                iCounter = iCounter + 1
                            End If
                        Next iClm
                    Next iRw
                Else
                    MsgBox "Incorrect calendar table format"
                End If
            End If
        Next tb
    End If
    Application.ScreenUpdating = True
End Sub

Public Function DaysOfMonth(dt As Date) As Integer
    If Month(dt) = 12 Then
        DaysOfMonth = 31
    Else
        DaysOfMonth = Day(DateSerial(Year(dt), Month(dt) + 1, 1) - 1)
    End If
End Function

Private Sub Document_Open()
    If UCase(Left(Application.System.OperatingSystem, 3)) = "MAC" Then
        MsgBox "This document is not suitable to be used on Mac"
    End If
End Sub

