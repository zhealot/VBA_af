VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Set ws = ThisWorkbook.Sheets("Ref")
    Set SCSheet = ThisWorkbook.Sheets("Shortcodes")
    Set InputSheet = ThisWorkbook.Sheets(3)
    Set RgSC = Nothing
    Set RgSiteBranch = Nothing
    Set RgSitePro = Nothing
    Set RgSiteSC = Nothing
    Set RgSite1 = Nothing
    Set RgSite2 = Nothing
    Set RgSite3 = Nothing
    Set RgLibrary = Nothing
    Set RgMeta = Nothing
    EnEvents = True

    blWork = False
    
End Sub
