VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Offer Letter"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "OAG_Offer of employment_UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub rbFixed_Click()
    Call FixPermanent
End Sub

Private Sub rbPermanent_Click()
    Call FixPermanent
End Sub

Sub FixPermanent()
    tbConcludeDate.Enabled = rbFixed.Value
    tbBusinessUnit.Enabled = Not rbFixed.Value
    tbBusinessUnit.Text = "Audit New Zealand"
    cbCredit.Enabled = Not rbFixed.Value
    cbSecurity.Enabled = Not rbFixed.Value
    Label12.Enabled = rbFixed.Value
    cbNolate.Enabled = Not rbFixed.Value
    Label12.Enabled = rbFixed.Value
    cbNolate.Enabled = Not rbFixed.Value
End Sub

Private Sub UserForm_Initialize()
    EnableControl fmPartTime, False
    Dim sHours(24) As String
    Dim i As Integer
    For i = 0 To 24
        cbHourFrom.AddItem i & ":00"
        cbHourFrom.AddItem i & ":30"
        cbHourTo.AddItem i & ":00"
        cbHourTo.AddItem i & ":30"
    Next
    tbDate.Text = Format(Now(), "dd-mm-yyyy")
    cbNolate.Value = False
    '############################################## for test
'    tbFirstName.Text = "John"
'    tbSurname.Text = "Doe"
'    tbCity.Text = "WLGN"
'    tbPostCode.Text = "3145"
'    tbSuburb.Text = "Subbbburbb"
'    tbStreetAddress.Text = "31415 road"
'    tbPosition.Text = "possistionnnn"
'    tbCommenceDate.Text = "14/3/2015"
'    tbLocation.Text = "Some where"
'    tbOfferCloseDate.Text = Date
'    tbRemuneration.Text = 63456

End Sub

Private Sub cbCancel_Click()
    Me.Hide
    ActiveDocument.Close False
End Sub

Private Sub cbOK_Click()
    Dim rg As Range
    If CheckTB(tbFirstName) Then
        Exit Sub
    End If
    If CheckTB(tbSurname) Then
        Exit Sub
    End If
    If CheckTB(tbStreetAddress) Then
        Exit Sub
    End If
    If CheckTB(tbSuburb) Then
        Exit Sub
    End If
    If CheckTB(tbCity) Then
        Exit Sub
    End If
    If CheckTB(tbPostCode) Then
        Exit Sub
    End If
    If CheckTB(tbDate) Then
        Exit Sub
    End If
    If Not IsDate(tbDate) Then
        tbDate.SetFocus
        MsgBox "Offer date format not valid."
        Exit Sub
    End If
    
    If rbFixed.Value = False And rbPermanent.Value = False Then
        MsgBox "Please choose an offter type."
        Exit Sub
    End If
    
    If Not IsDate(tbCommenceDate) Then
        tbCommenceDate.SetFocus
        MsgBox "Commence date format not valid."
        Exit Sub
    End If
    If rbFixed.Value Then
        If Not IsDate(tbConcludeDate.Text) Then
            tbConcludeDate.SetFocus
            MsgBox "Conclude date format not valid."
            Exit Sub
        End If
        If DateDiff("d", tbCommenceDate.Text, tbConcludeDate.Text) <= 0 Then
            MsgBox "Conclude date should be after commence date."
            Exit Sub
        End If
    End If
    If Not IsDate(tbOfferCloseDate) Then
        tbOfferCloseDate.SetFocus
        MsgBox "Offer close date format not valid."
        Exit Sub
    End If
    If CheckTB(tbPosition) Then
        Exit Sub
    End If
    If CheckTB(tbBusinessUnit) Then
        Exit Sub
    End If
    If CheckTB(tbCommenceDate) Then
        Exit Sub
    End If
    If rbFixed.Value Then
        If CheckTB(tbConcludeDate) Then
            Exit Sub
        End If
    End If
    
    If cbPartTime.Value Then
        If cbHourFrom.Value = "" Then
            cbHourFrom.SetFocus
            MsgBox "Please choose start time."
            Exit Sub
        End If
        If cbHourTo.Value = "" Then
            cbHourTo.SetFocus
            MsgBox "Please choose end time."
            Exit Sub
        End If
        If DateDiff("n", cbHourFrom.Text, cbHourTo.Text) < 0 Then
            If MsgBox("Start time late then finish time, are you sure?", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        If cbMon.Value Or cbTue.Value Or cbWed.Value Or cbThu.Value Or cbFri.Value Then
        Else
            MsgBox "Please choose work day(s) for the part time offer."
            Exit Sub
        End If
        If CheckTB(tbManager) Then
            Exit Sub
        End If
        If CheckTB(tbTitle) Then
            Exit Sub
        End If
    End If
    If CheckTB(tbLocation) Then
        Exit Sub
    End If
    If CheckTB(tbRemuneration) Then
        Exit Sub
    End If
    If CheckTB(tbOfferCloseDate) Then
        Exit Sub
    End If
    
    If Not IsNumeric(tbRemuneration.Text) Then
        MsgBox "Please enter a number for remuneration."
        tbRemuneration.SetFocus
        Exit Sub
    End If
    
    Me.Hide
    
    'set values
    SetBM "bmFirstName", tbFirstName.Text
    SetBM "bmFirstName2", tbFirstName.Text
    SetBM "bmFirstName3", tbFirstName.Text
    SetBM "bmFirstName4", tbFirstName.Text
    SetBM "bmSurname", tbSurname.Text
    SetBM "bmSurname2", tbSurname.Text
    SetBM "bmStreetAddress", tbStreetAddress.Text
    SetBM "bmSuburb", tbSuburb.Text
    SetBM "bmCity", tbCity.Text
    SetBM "bmPostcode", tbPostCode.Text
    SetBM "bmDate", tbDate.Text
    SetBM "bmPosition", tbPosition.Text
    SetBM "bmPosition2", tbPosition.Text
    SetBM "bmBusinessUnit", tbBusinessUnit.Text
    SetBM "bmBusinessUnit2", tbBusinessUnit.Text
    SetBM "bmLocation", tbLocation.Text
    SetBM "bmRemuneration", Replace(tbRemuneration.Text, "$", "")
    SetBM "bmOfferCloseDate", Format(tbOfferCloseDate, "Long Date")
    
    'Fixed term /Permanent offer
    If rbFixed.Value Then
        SetBM "bmStartDate", "Term of employment"
        SetBM "bmSalaryReview", ""
        HideBM "bmProMem", True
        HideBM "bmDriver", True
        HideBM "bmTOIL", False
        HideBM "bmCredit", True
        SetBM "bmFixedTerm", "fixed term "
        SetBM "bmPeriodOfEmployment", "The period of fixed term employment commences on " & Format(tbCommenceDate.Text, "Long Date") & " and concludes on " & Format(tbConcludeDate.Text, "Long Date") & "."
        HideBM "bmReason", False
    Else
        SetBM "bmStartDate", "Start Date"
        SetBM "bmSalaryReview", "Your salary will be reviewed annually. Currently this occurs in July."
        HideBM "bmProMem", False
        HideBM "bmDriver", False
        HideBM "bmTOIL", True
        HideBM "bmCredit", False
        If cbNolate.Value Then
            SetBM "bmPeriodOfEmployment", "Your employment will commence on a date to be mutually agreed, but no later than " & Format(tbCommenceDate.Text, "Long Date") & "."
        Else
            SetBM "bmPeriodOfEmployment", "Your employment will commence on " & Format(tbCommenceDate.Text, "Long Date") & "."
        End If
        SetBM "bmFixedTerm", ""
        Set rg = ActiveDocument.Bookmarks("bmReason").Range
        rg.SetRange rg.Start - 1, rg.End
        rg.Text = ""
    End If
    
    'Part time / full time offer
    If cbPartTime.Value Then
        Dim sWeekDays As String 'park time week days
        Dim sWeekHours  As String 'part time week hours
        Dim iDays As Integer
        If cbMon.Value And cbTue.Value And cbWed.Value And cbThu.Value And cbFri.Value Then
            sWeekDays = ", Monday to Friday."
            iDays = 5
        Else
            sWeekDays = IIf(cbMon.Value, ", Monday", "") & IIf(cbTue.Value, ", Tuesday", "") & IIf(cbWed.Value, ", Wednesday", "") & IIf(cbThu.Value, ", Thursday", "") & IIf(cbFri.Value, ", Friday", "") & "."
            If cbMon.Value Then
                iDays = iDays + 1
            End If
            If cbTue.Value Then
                iDays = iDays + 1
            End If
            If cbWed.Value Then
                iDays = iDays + 1
            End If
            If cbThu.Value Then
                iDays = iDays + 1
            End If
            If cbFri.Value Then
                iDays = iDays + 1
            End If
        End If
        sWeekHours = iDays * DateDiff("n", cbHourFrom.Text, cbHourTo.Text) / 60
        'days/hours
        SetBM "bmHours", "Your usual hours of work will be " & sWeekHours & " hours per week, to be worked between " & cbHourFrom.Text & " and " & cbHourTo.Text & sWeekDays & " Your actual work pattern will be agreed by your manager, " & tbManager.Text & ", " & tbTitle.Text & "."
        SetBM "bmProrated", ", pro-rated for the " & sWeekHours & " hours per week worked"
    Else
        SetBM "bmHours", "Your usual hours of work will be 40 hours per week, to be worked between 8.00am and 5.30pm, Monday to Friday."
        SetBM "bmProrated", ""
    End If
    
    'KiwiSaver only if fixed term more than 28 days
    If rbFixed.Value Then
        If DateDiff("d", tbCommenceDate.Text, tbConcludeDate.Text) > 28 Then
            HideBM "bmKiwiSaver", False
        Else
            HideBM "bmKiwiSaver", True
        End If
    End If
    'Medical insurance only if fixed term morethan 6 months
    If rbFixed.Value Then
        If DateDiff("m", tbCommenceDate.Text, tbConcludeDate.Text) > 6 Then
            HideBM "bmMedical", False
        Else
            HideBM "bmMedical", True
        End If
    End If

    
    If rbPermanent.Value Then
        'credit history
        If cbCredit.Value Then
            SetBM "bmCreditHistory", " and credit"
            SetBM "bmCreditHistory2", " and credit"
            SetBM "bmCreditHistory3", "s"
            HideBM "bmCreditClause", False
        Else
            SetBM "bmCreditHistory", ""
            SetBM "bmCreditHistory2", ""
            SetBM "bmCreditHistory3", ""
            HideBM "bmCreditClause", True
        End If
        'security
        If cbSecurity.Value Then
            HideBM "bmSecurityYes", False
            Set rg = ActiveDocument.Bookmarks("bmSecurityNo").Range
            rg.SetRange rg.Start - 1, rg.End
            rg.Text = ""
        Else
            HideBM "bmSecurityYes", True
            HideBM "bmSecurityNo", False
        End If
    Else
        HideBM "bmSecurityYes", True
        HideBM "bmSecurityNo", False
        HideBM "bmCreditClause", True
        
    End If
  
    SetBM "bmOfferCloseDate", Format(tbOfferCloseDate.Value, "Long Date")
    'lock down document
    'ActiveDocument.Protect wdAllowOnlyReading, , "7060"
End Sub

'replace bookmark content text
Function SetBM(bm As String, txt As String)
    Dim rg As Range
    If ActiveDocument.Bookmarks.Exists(bm) Then
        Set rg = ActiveDocument.Bookmarks(bm).Range
        rg.Text = (txt)
        ActiveDocument.Bookmarks.Add bm, rg
    End If
End Function

'hide/show bookmark content
Function HideBM(bm As String, YesNo As Boolean)
    If ActiveDocument.Bookmarks.Exists(bm) Then
        ActiveDocument.Bookmarks(bm).Range.Font.Hidden = YesNo
    End If
End Function

'check whether textbox filled
Function CheckTB(tb As TextBox) As Boolean
    If Trim(tb.Text) = "" Then
        CheckTB = True
        tb.SetFocus
        MsgBox "Please fill this filed."
    End If
End Function

Function EnableControl(fm As MSForms.Frame, YesNo As Boolean)
    Dim ctr As Control
    fm.Enabled = YesNo
    For Each ctr In fm.Controls
        ctr.Enabled = YesNo
    Next ctr
End Function

Private Sub cbPartTime_Click()
    EnableControl fmPartTime, cbPartTime.Value
End Sub
