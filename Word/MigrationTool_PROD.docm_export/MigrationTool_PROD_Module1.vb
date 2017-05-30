Attribute VB_Name = "Module1"
Public Enum JohnsCustomBeforeInsideOrAfter
    jcBefore = -1
    jcInside = 0
    jcAfter = 1
End Enum

Public Function Log(str As String) As Boolean
    Open ThisDocument.Path & "\log.txt" For Append As #1
    Print #1, Now & ":" & vbTab & str
    Close #1
End Function

Public Function ToITemplate(ToI As String) As String
    Select Case Trim(ToI)
        Case "Code of Welfare"
            ToITemplate = "MPI_CodeOfWelfare_Template.docx"
        Case "Operational Code"
            ToITemplate = "MPI_OPCODE_Template.docx"
        Case "Guidance Document"
            ToITemplate = "MPI_Guidance_Template.docx"
        Case "Plant Export Requirement"
            ToITemplate = "MPI_PE_Template.docx"
'        Case "MPI_FYI_Template"
'            ToITemplate = "MPI_FYI_Template.docx"
        Case "Wine Notice"
            ToITemplate = "MPI_Wine_Notice_Template.docx"
'        Case "Wine Notice"
'            ToITemplate = "MPI_Wine_OMAR_Template.docx"
        Case "Plant Import Requirement"
            ToITemplate = "MPI_PI_Template.docx"
        Case "Organic Export Requirement"
            ToITemplate = "MPI_ORG_Template.docx"
        Case "Import Health Standard"
            ToITemplate = "MPI_IHS_Template.docx"
        Case "Food Standard"
            ToITemplate = "MPI_FOOD_STD_Template.docx"
        Case "Food Notice"
            ToITemplate = "MPI_FOOD_Notice_Template.docx"
        Case "Animal Products Notice"
            ToITemplate = "MPI_AP_Notice_Template.docx"
        Case "Facility Standard"
            ToITemplate = "MPI_FAC_Template.docx"
        Case "Craft Risk Management Standard"
            ToITemplate = "MPI_CRM_Template.docx"
        Case "ACVM Requirement"
            ToITemplate = "MPI_ACVM_Req_Template.docx"
        Case "Treatment Requirement"
            ToITemplate = "MPI_TREAT_Template.docx"
        Case "ACVM Notice"
            ToITemplate = "MPI_ACVM_Notice_Template.docx"
        Case Else
            ToITemplate = ""
    End Select
End Function

Function PopulateArray(ByRef Ar() As String, key As String)
    Ar(0) = key
    Select Case key
        Case "For Your Information"
            Ar(1) = "08be437b-da17-4333-83bc-962f4002d684"
        Case "Guidance Document"
            Ar(1) = "c4a8b94d-4b6b-4094-9d7c-4dfebe6bee0d"
        Case "Other Guidance"
            Ar(1) = "7f603f21-e7cc-4b67-9bd7-275034fdd5b8"
        Case "ACVM Notice"
            Ar(1) = "8e3bd1ee-0171-4deb-a632-09e77c654e31"
        Case "ACVM Requirement"
            Ar(1) = "ac282ebc-548b-4f43-835e-1fe3951ac271"
        Case "Animal Products Notice"
            Ar(1) = "c3e339b3-a8ca-4a04-a34a-a2e11bb74003"
        Case "Code of Welfare"
            Ar(1) = "17bc35ce-358f-42ee-9e45-e7e4d242898a"
        Case "Craft Risk Management Standard"
            Ar(1) = "5f49d691-2dd3-46bb-8022-fbdb15452d88"
        Case "Facility Standard"
            Ar(1) = "95c29740-b330-4b1d-b939-d497fd9bea85"
        Case "Food Notice"
            Ar(1) = "b74dce2e-0dd5-4742-8251-27360687ebae"
        Case "Food Standard"
            Ar(1) = "cfc8947e-e2d7-4061-a19d-8967570cf7b3"
        Case "Import Health Standard"
            Ar(1) = "99168b8a-db5f-446b-b56e-afdaefbd5b0d"
        Case "Operational Code"
            Ar(1) = "1b5c0c82-4712-4fe5-8303-5db017cf95c8"
        Case "Organic Export Requirement"
            Ar(1) = "0d717db3-2af5-489f-8f58-0c7fe2c4f54f"
        Case "Plant Export Requirement"
            Ar(1) = "b437ee44-d459-4e80-baca-d0ff649b8b21"
        Case "Plant Import Requirement"
            Ar(1) = "7011ba94-1b25-41f7-a3b4-235e9fd0623f"
        Case "Treatment Requirement"
            Ar(1) = "dbbdb82c-3f46-4aac-90b1-88b6e63ba64e"
        Case "Wine Notice"
            Ar(1) = "04ae4fff-8ca5-49d8-a8a3-edcd10851aa0"
        Case Else
            Ar(0) = ""
            Ar(1) = ""
        End Select
End Function

Sub ToggleContentControlPresentOrAbsent( _
    doc As Word.Document _
    , tagOfContentControl As String _
    , isControlPresentOrAbsent As Boolean _
    , tagOfContentControl2insertAfterOrInside As String _
    , insertBeforeInsideOrAfter As JohnsCustomBeforeInsideOrAfter _
)
    Dim ccs As Word.ContentControls
    Dim cc As Word.ContentControl
    Dim rng As Word.Range
    Dim pos As JohnsCustomBeforeInsideOrAfter
    pos = insertBeforeInsideOrAfter
    On Error Resume Next
    Set ccs = doc.SelectContentControlsByTag(tagOfContentControl)
    Set cc = ccs(1)
    Set rng = cc.Range
    If Err.Number = 0 Then
        If isControlPresentOrAbsent Then
            MoveContentControlIntoQuickpart cc
        End If
    Else
        Err.Clear
        If Not isControlPresentOrAbsent Then
            InsertContentControlRelativeToOtherCc _
                wordDocument:=doc _
                , tagOfContentControl2insertInside:=tagOfContentControl2insertAfterOrInside _
                , tagOfContentControl2getFromAutotext:=tagOfContentControl _
                , insertBeforeInsideOrAfter:=pos
        End If
    End If
    
    On Error GoTo 0
End Sub

Sub MoveContentControlIntoQuickpart(cc As Word.ContentControl)
    Dim autoTxt As Word.AutoTextEntry
    Dim tagOfContentControl As String
    Dim rng As Word.Range
    
    tagOfContentControl = cc.Tag
    
    With Word.ActiveDocument
        For Each autoTxt In .AttachedTemplate.AutoTextEntries
            If LCase(autoTxt.Name) = LCase(tagOfContentControl) Then
                GoTo JustRemoveObjectsFromDoc
            End If
        Next autoTxt
        .AttachedTemplate.AutoTextEntries.Add _
            Name:=tagOfContentControl _
            , Range:=cc.Range
JustRemoveObjectsFromDoc:
        Set rng = cc.Range
        cc.LockContentControl = False
        cc.Delete DeleteContents:=True
        rng.Delete
    End With
End Sub

Sub InsertContentControlRelativeToOtherCc( _
    wordDocument As Word.Document _
    , tagOfContentControl2insertInside As String _
    , tagOfContentControl2getFromAutotext As String _
    , Optional insertBeforeInsideOrAfter As JohnsCustomBeforeInsideOrAfter _
)
    Dim ccs As Word.ContentControls
    Dim cc As Word.ContentControl
    Dim rng As Word.Range
    
    With wordDocument
        Set ccs = .SelectContentControlsByTag(tagOfContentControl2insertInside)
        If ccs Is Nothing Then
            Exit Sub
        End If
        
        Set cc = ccs(1)
        Set rng = cc.Range
        
        If insertBeforeInsideOrAfter = jcBefore Then
            rng.MoveStart Unit:=wdCharacter, Count:=-1
            rng.Collapse Word.wdCollapseStart
        ElseIf insertBeforeInsideOrAfter = jcAfter Then
            rng.MoveEnd Unit:=wdCharacter, Count:=2
        End If
        
        If insertBeforeInsideOrAfter <> jcBefore Then
            rng.Collapse Word.wdCollapseEnd
        End If
        
        .AttachedTemplate.AutoTextEntries(tagOfContentControl2getFromAutotext).Insert _
            Where:=rng _
            , RichText:=True
        On Error GoTo Catch
        
        ActiveDocument.Bookmarks("bkmUnprotect").Range.Editors.Add wdEditorEveryone
        ActiveDocument.Bookmarks("bkmUnprotect").Delete
        
Catch:
        On Error GoTo 0
    End With
End Sub

Function contentControlExists(cc As String) As Boolean
    Dim exists As Boolean
    exists = False
    
    Dim controls As ContentControls
    Set controls = ActiveDocument.SelectContentControlsByTag(cc)
    If controls.Count > 0 Then
        exists = True
    End If
    contentControlExists = exists
End Function

Sub export()
    Dim i As Integer
    For i = 1 To ThisDocument.VBProject.VBComponents.Count
        ThisDocument.VBProject.VBComponents(1).export ThisDocument.Path & "\" & ThisDocument.VBProject.VBComponents(i).Name & "_code.vb"
    Next i
    
End Sub
