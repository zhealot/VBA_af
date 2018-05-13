VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Logos"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   OleObjectBlob   =   "Ministerials_Briefing Template_UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbBio_Click()
    Call CheckBoxEvent(cbBio)
End Sub

Private Sub cbFis_Click()
    Call CheckBoxEvent(cbFis)
End Sub

Private Sub cbFor_Click()
    Call CheckBoxEvent(cbFor)
End Sub

Private Sub cbMPI_Click()
    Call CheckBoxEvent(cbMPI)
End Sub


Private Sub cbNZF_Click()
    Call CheckBoxEvent(cbNZF)
End Sub

Private Sub cbO_Click()
    If cbO.Value Then
        If sOthers <> " " Then
            cbOthers.Value = sOthers
        End If
    Else
        sOthers = " "
    End If
    cbOthers.Visible = cbO.Value
    cbOthers.Enabled = cbO.Value
    fmOthers.Visible = cbO.Value
End Sub

Private Sub cbOthers_Change()
    Dim ctr As control
    For Each ctr In fmOthers.Controls
        If ctr.Visible = True Then ctr.Visible = False
        If Right(ctr.Name, Len(ctr.Name) - 3) = LCase(cbOthers.Value) Then
            ctr.Visible = True
            sOthers = cbOthers.Value
        End If
    Next
End Sub

Private Sub cmbCancel_Click()
    Unload Me
End Sub

Private Sub cmbOK_Click()
    Unload Me
    Dim doc As Document
    Dim rg As Range
    Dim sp As InlineShape
    Dim HasLogo As Boolean 'row has any logo shown
    Set doc = ActiveDocument
    'save selected value to custom document properties
    doc.CustomDocumentProperties(PROPERTYNAME).Value = IIf(sLogos = "", " ", sLogos)
    doc.CustomDocumentProperties("others").Value = IIf(sOthers = "", " ", sOthers)
    'hide/show logos

    Dim tmp As Template
    For Each tmp In Application.Templates
        If tmp.Name = ThisDocument.Name Then
            'in top row
            Dim iCounter As Integer
            iCounter = 3 - Len(Replace(Replace(sLogos, "Fis", ""), "For", "")) / 3  'factor to adjust image width in top row
            Set rg = doc.Bookmarks(LOGOBOOKMARK).Range
            If rg.Cells.Count > 0 Then
                Set rg = rg.Cells(1).Range
                'delete existing logos
                If rg.InlineShapes.Count > 0 Then
                    For i = rg.InlineShapes.Count To 1 Step -1
                        rg.InlineShapes(i).Delete
                    Next
                End If
                rg.Collapse wdCollapseEnd
            Else
                Exit Sub
            End If
            'insert first row logos
            Set rg = doc.Bookmarks(LOGOBOOKMARK).Range
            If InStr(sLogos, "Bio") > 0 Then
                Set rg = tmp.BuildingBlockEntries("Bio").Insert(rg)
                rg.Collapse wdCollapseEnd
            End If
            If InStr(sLogos, "MPI") > 0 Then
                Set rg = tmp.BuildingBlockEntries("MPI").Insert(rg)
                rg.Collapse wdCollapseEnd
            End If
            If InStr(sLogos, "NZF") > 0 Then
                Set rg = tmp.BuildingBlockEntries("NZF").Insert(rg)
                rg.Collapse wdCollapseEnd
            End If
            Set rg = doc.Bookmarks(LOGOBOOKMARK).Range.Cells(1).Range
            If rg.InlineShapes.Count > 0 Then
                HasLogo = False
                For Each sp In rg.InlineShapes
                    HasLogo = True
                    sp.LockAspectRatio = msoTrue
                    sp.Width = CentimetersToPoints(5.4 + iCounter * 1.6)
                Next sp
            End If
            
            'in bottom row
            Set rg = doc.Bookmarks(LOGOBOOKMARK2).Range
            If rg.Cells.Count > 0 Then
                Set rg = rg.Cells(1).Range
                'delete existing logos
                If rg.InlineShapes.Count > 0 Then
                    For i = rg.InlineShapes.Count To 1 Step -1
                        rg.InlineShapes(i).Delete
                    Next
                End If
                rg.Collapse wdCollapseEnd
            Else
                Exit Sub
            End If
            iCounter = IIf(sOthers = " ", 3, 2) - Len(Replace(Replace(Replace(sLogos, "MPI", ""), "NZF", ""), "Bio", "")) / 3 'factor to adjust image width in bottom row
            'insert other logo
            If sOthers <> " " And sOthers <> "" Then
                Set rg = doc.Bookmarks(LOGOBOOKMARK2).Range
                rg.Collapse wdCollapseEnd
                Set rg = tmp.BuildingBlockEntries(sOthers).Insert(rg)
                rg.Collapse wdCollapseEnd
                HasLogo = True
            End If
            Set rg = doc.Bookmarks(LOGOBOOKMARK2).Range
            rg.Collapse wdCollapseEnd
            If InStr(sLogos, "Fis") > 0 Then
                Set rg = tmp.BuildingBlockEntries("Fis").Insert(rg)
                rg.Collapse wdCollapseEnd
            End If
            If InStr(sLogos, "For") > 0 Then
                Set rg = tmp.BuildingBlockEntries("For").Insert(rg)
                rg.Collapse wdCollapseEnd
            End If
            Set rg = doc.Bookmarks(LOGOBOOKMARK2).Range.Cells(1).Range
            If rg.InlineShapes.Count > 0 Then
                HasLogo = False
                For Each sp In rg.InlineShapes
                    'delete other logo first
                    If sp.Title = "others" Then
                        sp.Delete
                    ElseIf sp.Title <> "" Or sp.Title <> " " Then
                        sp.Range.Font.Hidden = IIf(InStr(sLogos, sp.Title) > 0, False, True)
                        If sp.Range.Font.Hidden = False Then
                            sp.LockAspectRatio = msoTrue
                            sp.Width = CentimetersToPoints(5.4 + 1.6 * iCounter)
                        End If
                        If sp.Range.Font.Hidden = False Then
                            HasLogo = True
                        End If
                    End If
                Next sp
            End If
        End If
    Next tmp

'    If Not HasLogo Then
'        doc.Bookmarks(LOGOBOOKMARK2).Range.Cells(1).Range.Font.Hidden = True
'        doc.Bookmarks(LOGOBOOKMARK2).Range.Cells(1).Height = InchesToPoints(ROW_HIDE)
'    End If
End Sub


Private Sub UserForm_Initialize()
    On Error Resume Next
    'get logos
    sLogos = ActiveDocument.CustomDocumentProperties(PROPERTYNAME)
    sOthers = ActiveDocument.CustomDocumentProperties("others")
    SwitchLogos "Bio"
    SwitchLogos "Fis"
    SwitchLogos "For"
    SwitchLogos "MPI"
    SwitchLogos "NZF"
    cbO.Value = IIf(sOthers = " ", False, True)
    If cbO.Value Then
        cbOthers.Value = sOthers
    End If
    'populate combo box
    Dim BBCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    BBCount = ActiveDocument.AttachedTemplate.BuildingBlockEntries.Count
    Dim entries(1 To LOGOCOUNT, 1 To 2) As String
    For i = 1 To BBCount
        entries(i, 2) = ActiveDocument.AttachedTemplate.BuildingBlockEntries(i).Name
        entries(i, 1) = ActiveDocument.AttachedTemplate.BuildingBlockEntries(i).Description
    Next i
    'sort array
    Dim aTmp(1 To 1, 1 To 2) As String
    For i = 1 To LOGOCOUNT
        For j = i To LOGOCOUNT
            If entries(i, 1) > entries(j, 1) Then
                aTmp(1, 1) = entries(i, 1)
                aTmp(1, 2) = entries(i, 2)
                entries(i, 1) = entries(j, 1)
                entries(i, 2) = entries(j, 2)
                entries(j, 1) = aTmp(1, 1)
                entries(j, 2) = aTmp(1, 2)
            End If
        Next j
    Next i
    'assign value to combo box
    cbOthers.List = entries
    'load images
    imginz.Picture = LoadPicture("D:\Temp\image1.jpeg")
End Sub

'tick/untick checkbox based on checkbox name
Function SwitchLogos(s As String)
    Me.Controls("cb" & s).Value = IIf(InStr(sLogos, s) > 0, True, False)
    CheckBoxEvent Me.Controls("cb" & s)
End Function
