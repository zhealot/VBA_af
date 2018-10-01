VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "NEO Proposal Template"
   ClientHeight    =   6900
   ClientLeft      =   180
   ClientTop       =   705
   ClientWidth     =   8295
   OleObjectBlob   =   "NEO_Allfields_UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------   -----------------------------------------
' Developed for NEO (Ergo)
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             September 2018
' Description:      populate job title etc, fix blank page issue.
'------------------------------------------------------------------------------

Private Sub cbCover_Click()
    cbCoverImage.Enabled = cbCover.Value
End Sub

Private Sub cbRevision_Click()
    Me.Label4.Enabled = cbRevision.Value
    Me.tbRevision.Enabled = cbRevision.Value
End Sub

'Allfields Customised Solutions 25/10/2016

Private Sub CommandButton1_Click()
    'validate title/client/author
    If Not CheckTB(tbJobTitle, "Job Title.") Then Exit Sub
    If Not CheckTB(tbTitle, "Title of document.") Then Exit Sub
    If Not CheckTB(tbClient, "Client name.") Then Exit Sub
    If Not CheckTB(tbDocumentNumber, "Document Number.") Then Exit Sub
    If Not CheckTB(tbAuthor, "Author name.") Then Exit Sub
    If cbRevision.Value Then
        If Not CheckTB(tbRevision, "Revision.") Then Exit Sub
    End If
    'validate date
    If Not IsDate(tbDate.Text) Then
        tbDate.SetFocus
        MsgBox "Please enter a valid date."
        Exit Sub
    End If
                    
    If cbCover.Value Then
        If Trim(cbCoverImage.Caption) = "" Or Dir(cbCoverImage.Caption) = "" Then
            MsgBox "Please select a cover image."
            Exit Sub
        End If
    End If
    
    Me.Hide
    Dim doc As Document
    Set doc = ActiveDocument
    Dim sp As Shape
    Dim rg As Range
    
    'insert cover image
    If cbCover.Value Then
        If cbCoverImage.Caption <> "" Then
            'delete existing cover image but keep the 'rings'
            For Each sp In doc.Shapes
                If sp.AlternativeText = "cove image" Then
                    sp.Delete
                End If
            Next sp
            Set rg = doc.Content
            rg.Collapse wdCollapseStart
            Dim insp As InlineShape
            
            Set insp = rg.InlineShapes.AddPicture(cbCoverImage.Caption)
            Set sp = insp.ConvertToShape
            With sp
                .WrapFormat.Type = wdWrapBehind
                .ZOrder msoSendToBack
                .ZOrder msoBringForward
                .LockAspectRatio = msoFalse
                .Width = 548.79
                .Height = .Width * 0.7 '(3 / 4) 'set aspect ratio to 4:3
                .RelativeHorizontalPosition = 1
                .RelativeVerticalPosition = 1
                .Left = 22.96
                .Top = 107
                .AlternativeText = "cove image"
                'sp.Fill.PictureEffects.Insert(msoEffectSharpenSoften).EffectParameters(1).Value = -0.4
            End With
        End If
    End If
    
    'write back variables
    doc.CustomDocumentProperties("JobTitle").Value = tbJobTitle.Text
    doc.BuiltInDocumentProperties("Title").Value = tbTitle.Text
    doc.CustomDocumentProperties("Client").Value = tbClient.Text
    doc.CustomDocumentProperties("DocumentNumber").Value = tbDocumentNumber.Text
    doc.CustomDocumentProperties("Date").Value = Format(tbDate.Text, DateFormat)
    doc.CustomDocumentProperties("Author").Value = tbAuthor.Text
    doc.CustomDocumentProperties("Revision").Value = tbRevision.Text
    'doc.CustomDocumentProperties("Project").Value = tbProject.Text
    doc.CustomDocumentProperties("IsNew").Value = "No" 'mark as a edited document
        
    'write new revision
    If cbRevision.Value Then
        AddRevision
    End If
    'updates cc in content
    UpdateCC
    
    'prepare PDF version by change all sectionbreaks into 'Next Page' instead of 'Odd/Even' page, store those pages for restoring
    Dim i As Integer
    Dim sTmp As String
    If cbPDF.Value Then
        For i = 1 To doc.Sections.Count
            'so the value for each section is: ','+'SectionStartType'
            sTmp = sTmp & "," & doc.Sections(i).PageSetup.SectionStart
            'ugly fix to workaround Word's restriction on blank pages for new, page-number-restarted section
            If i = 4 Then
                doc.Sections(i).PageSetup.SectionStart = wdSectionContinuous
                Set rg = doc.Sections(i).Range
                rg.SetRange rg.Start, rg.End - 1
                rg.Text = ""
            Else
                doc.Sections(i).PageSetup.SectionStart = wdSectionNewPage
            End If
        Next i
        'store section page setup info to document property
        doc.CustomDocumentProperties("SecPageNo").Value = sTmp
    Else
        'fix blank page issue prior of content
        If doc.Sections.Count > 4 Then
            'blank page section is deleted, need re-insert
            Dim rg4 As Range
            Dim rg3 As Range
            Set rg4 = doc.Sections(4).Range
            rg4.Collapse wdCollapseEnd
            Set rg3 = doc.Sections(3).Range
            If rg4.Information(wdActiveEndPageNumber) - rg3.Information(wdActiveEndPageNumber) > 3 Then
                Set rg = doc.Sections(3).Range
                rg.Collapse wdCollapseEnd
                rg.InsertBreak wdSectionBreakNextPage
                Set rg = doc.Sections(4).Range
                rg.Style = wdStyleNormal
                rg.SetRange rg.Start, rg.End - 1
                If doc.Sections(3).Range.Information(wdActiveEndPageNumber) Mod 2 = 1 Then
                    doc.Sections(4).PageSetup.SectionStart = wdSectionEvenPage
                    rg.Text = "Please DO NOT delete this blank page."
                Else
                    doc.Sections(4).PageSetup.SectionStart = wdSectionContinuous
                    rg.Text = ""
                End If
                doc.Sections(5).PageSetup.SectionStart = wdSectionNewPage
                On Error Resume Next
                For i = 1 To 3
                    doc.Sections(5).Footers(i).LinkToPrevious = False
                    doc.Sections(5).Headers(i).LinkToPrevious = False
                Next i
                For i = 1 To 3
                    doc.Sections(4).Footers(i).LinkToPrevious = False
                    doc.Sections(4).Footers(i).Range.Delete
                    doc.Sections(4).Headers(i).LinkToPrevious = False
                    doc.Sections(4).Headers(i).Range.Delete
                Next i
            End If
        End If
        sTmp = Trim(doc.CustomDocumentProperties("SecPageNo").Value)
        If sTmp <> "" Then  'has records of sections
            Dim ary() As String
            ary = Split(sTmp, ",")
            'NOTE: ary has one heading ',' in it
            If UBound(ary) = doc.Sections.Count Then    'make sure stored value has same number of sections as document does
                For i = 1 To doc.Sections.Count
                    'ugly fix to make content starts from odd page, if necessary add a blank page with warning.
                    If i = 4 Then
                        Set rg = doc.Sections(i).Range
                        rg.SetRange rg.Start, rg.End - 1    'exclude section break mark
                        If doc.Sections(3).Range.Information(wdActiveEndPageNumber) Mod 2 = 1 Then  'page prior content on odd page, need blank page inserted
                            doc.Sections(i).PageSetup.SectionStart = wdSectionEvenPage
                            rg.Text = "Please DO NOT delete this blank page"
                        Else
                            doc.Sections(i).PageSetup.SectionStart = wdSectionContinuous
                            rg.Text = ""
                        End If
                    Else
                        doc.Sections(i).PageSetup.SectionStart = ary(i)
                    End If
                Next i
            Else
                MsgBox "Document has different sections than before, unable to restore blank pages."
            End If
        Else
            Set rg = doc.Sections(4).Range
            rg.Style = wdStyleNormal
            rg.SetRange rg.Start, rg.End - 1    'exclude section break mark
            If doc.Sections(3).Range.Information(wdActiveEndPageNumber) Mod 2 = 1 Then  'page prior content on odd page, need blank page inserted
                doc.Sections(4).PageSetup.SectionStart = wdSectionEvenPage
                rg.Text = "Please DO NOT delete this blank page"
            Else
                doc.Sections(4).PageSetup.SectionStart = wdSectionContinuous
                rg.Text = ""
            End If
        End If
        doc.CustomDocumentProperties("SecPageNo").Value = " "
    End If
End Sub


Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub cbCoverImage_Click()
    Dim oDlg As Dialog
    Set oDlg = Application.Dialogs(wdDialogInsertPicture)
    oDlg.Display
    If oDlg.Name <> "" Then
        cbCoverImage.Caption = oDlg.Name
    Else
        cbCoverImage.Caption = "Click to select cover image"
    End If
End Sub

Sub UserForm_Initialize()
    If ThisDocument.CustomDocumentProperties("IsNew").Value = "Yes" Then
        FormInit True
    Else
        FormInit False
    End If
End Sub
