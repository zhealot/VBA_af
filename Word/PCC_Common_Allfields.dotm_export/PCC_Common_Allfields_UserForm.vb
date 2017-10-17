Attribute VB_Name = "UserForm"
Sub UserSetup(control As IRibbonControl)
'   Loads the UserForm1 and displays it
Load UserForm1
UserForm1.Show

End Sub

Sub Autoexec()

txtUser = System.PrivateProfileString(strDefaultUserIni, "UserSetup", "User")
If txtUser = "" Then
MsgBox "Please take the time to enter in some default information that will be used in a number of the templates.", vbInformation, "User Setup Information"
Load UserForm1
UserForm1.Show
End If
   
End Sub

Public Sub ReplaceTextInSelection(sFind As String, sReplace As String)
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = sFind
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = sFind
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub

'********************************************************************

' Function :            RestorePosition
'
' Arguments         None

' Returns               - 1(TRUE) position restored
'                       0(FALSE) position could Not be restored

' Description:  This Function set the cursor position To the position  stored in the bookmark "origin" set by save position .  If the bookmark does Not exist Or the position could Not be restored For Any reason the Function returns FALSE .  The "origin" boo
'kmark Is deleted by the Function .

'********************************************************************

Public Function RestorePosition() As Boolean

    Dim intRetVal As Boolean

    intRetVal = False

    On Error GoTo ExitRestorePosition

    If ActiveDocument.Bookmarks.Exists("origin") Then
        ActiveDocument.Bookmarks("origin").Select
        ActiveDocument.Bookmarks("origin").Delete
        intRetVal = True
    End If

ExitRestorePosition:
    RestorePosition = intRetVal

End Function

'********************************************************************

' Sub :         SavePosition

' Description:  This subroutine saves the current cursor position by creating a bookmark called origin .

'********************************************************************

Public Sub SavePosition()

    ActiveDocument.Bookmarks.Add Name:="origin", Range:=Selection.Range

End Sub

Public Sub ReplaceBookmarkText(sBookmark As String, sText As String)
'********************************************************************
' Sub :         ReplaceBookmarkText(sBookmark, sText)

' Arguments     sBookmark - The name of the bookmark to goto
'               sText - The Text To Insert at the bookmark

' Description:  This sub replaces the text of a certain bookmark with new text.

'********************************************************************

    Dim StartPos
    Dim EndPos

    ' Check the bookmark exists
    'If Not BookmarkExists(sBookmark) Then Exit Sub

'   Clear the text in the original bookmark
    ActiveDocument.Bookmarks(sBookmark).Select
    If Selection.Start <> Selection.End Then
        Selection.Range.Delete
    End If
    StartPos = Selection.Start
    Selection.TypeText sText

    ' Select the text we have just inserted and name the bookmark
    EndPos = Selection.Start
    ActiveDocument.Range(StartPos, EndPos).Select
    ActiveDocument.Bookmarks.Add Name:=sBookmark, Range:=Selection.Range

End Sub

'Sub open_QAS()
'Shell ("C:\Program Files\QAS\QuickAddress Pro\QAPrown.exe")
'Shell "wscript.exe W:\!Common\Templates\QAS.vbs"
'End Sub
    
'Sub Powerpoint()
'Shell "Wscript.exe W:\!Common\Templates\Allfields_Setup\powerpoint.vbs"
'End Sub
