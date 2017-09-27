ThisDocument.AttachedTemplate.BuildingBlockEntries.Add Name:=Left(bbName, 32), _
													  Type:=wdTypeQuickParts, _
													  Category:="General", _
													  Description:=bbName, _
													  Range:=rg, _
													  InsertOptions:=wdInsertContent

doc.AttachedTemplate.BuildingBlockEntries(Blocks(i).Name).Insert rg, True

    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select Autotext document."
        .Filters.Clear
        .Filters.Add "Word document", "*.docx", 1
        .InitialFileName = ThisDocument.Path & "\"
        .InitialView = msoFileDialogViewDetails
        .ButtonName = "OK"
        If .Show = -1 Then
            sFilename = .SelectedItems(1)
        End If
    End With

