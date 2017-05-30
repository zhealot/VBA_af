VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_New()
    Load frmDocProp
    frmDocProp.Show vbModal
    
    AutoCaptions("Microsoft Word Table").AutoInsert = True

'   Open the Document at the Bookmarked Area
    
    If ActiveDocument.Bookmarks.Exists("StartLocation") = True Then
    ActiveDocument.Bookmarks("StartLocation").Select
    
    End If
    
'   Display Styles Window and hide Protection Window
    
    On Error Resume Next
    Application.TaskPanes(wdTaskPaneFormatting).Visible = True
    Application.TaskPanes(wdTaskPaneDocumentProtection).Visible = False
    'ActiveWindow.View.ShadeEditableRanges = False
    
'    ActiveDocument.TrackRevisions = True
'    ActiveDocument.ShowRevisions = True
    
End Sub

Private Sub Document_Open()
   
'   Open the Document at the Bookmarked Area
    
    If ActiveDocument.Bookmarks.Exists("StartLocation") = True Then
    ActiveDocument.Bookmarks("StartLocation").Select

    End If
    
'   update all fields on open, tao@allfields.co.nz, 6/11/2015
    Call UpdateHeadersFooters

    On Error Resume Next
    Application.TaskPanes(wdTaskPaneFormatting).Visible = True
    Application.TaskPanes(wdTaskPaneDocumentProtection).Visible = False
    ActiveWindow.View.ShadeEditableRanges = False
        
End Sub


Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl)

'   Insert an image border on Content Controls

    Call ImgBorder

End Sub


