Attribute VB_Name = "Module1"
'----------------------------------   -------------------------------------------
' Developed for Ministry for Primary Industries
' Created by:       Allfields Customised Solutions Limited
' Contact Info:     hello@allfields.co.nz, 04 978 7101
' Date:             April 2018
' Description:      hide/show logos based on choice
'-----------------------------------------------------------------------------
Public Const LOGOCOUNT = 30    'number of "other" logos
Public Const LOGOBOOKMARK = "bmLogos" 'bookmark of the logos cell
Public Const LOGOBOOKMARK2 = "bmLogos2"
Public Const PROPERTYNAME = "logos"
Public Const ROW_SHOW = 1
Public Const ROW_HIDE = 0.1
Public sLogos As String 'Bio/Fis/For/MPI/NZF
Public sOthers As String

Sub Callback(control As IRibbonControl)
    UserForm1.Show
End Sub


Public Function CheckBoxEvent(cb As MSForms.CheckBox)
    Dim s As String
    s = Right(cb.Name, Len(cb.Name) - 2)
    'set image to show/hide
    UserForm1.Controls.Item("img" & s).Visible = cb.Value
    'keep value of selected logos
    If cb.Value Then    'ticked on
        If InStr(sLogos, s) <= 0 Then   'value not in sLogos
            sLogos = sLogos & s
        End If
    Else        'ticked off
        sLogos = Replace(sLogos, s, "")
    End If
End Function

Sub AddBuildingBlocks()
    Dim docS As Document
    Set docS = Documents.Open("D:\Box Sync\2. Staff Related Activities\Tao's playground\WorkingInProgress\MPI_WIP\Logos.docx", False, True, , , , , , , , , False)
    Dim rg As Range
    Dim tb As Table
    Dim sp As InlineShape
    Dim sName As String
    Dim sDcrp As String
    
    Set tb = docS.Tables(1)
    Dim iRow As Integer
    For i = 1 To tb.Rows.Count
        Set rg = tb.Cell(i, 2).Range
        If rg.Paragraphs.Count > 1 Then
            sDcrp = Left(rg.Paragraphs(1).Range.Text, Len(rg.Paragraphs(1).Range.Text) - 1)
            sName = Left(rg.Paragraphs(2).Range.Text, Len(rg.Paragraphs(2).Range.Text) - 2)
        End If
        Set rg = tb.Cell(i, 1).Range
        If rg.InlineShapes.Count > 0 Then
            Set sp = rg.InlineShapes(1)
            Set rg = sp.Range
            ThisDocument.AttachedTemplate.BuildingBlockEntries.Add Name:=sName, Type:=wdTypeQuickParts, Category:="General", Description:=sDcrp, Range:=rg, InsertOptions:=wdInsertContent
        End If
    Next i
    docS.Close False
    Set docS = Nothing
End Sub
