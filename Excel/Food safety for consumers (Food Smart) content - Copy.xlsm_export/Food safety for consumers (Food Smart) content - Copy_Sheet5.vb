VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SaveAll()
'    Dim fsoFSO As Scripting.FileSystemObject
    Dim cll As Range
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim sUrl As String
    Dim ext As String
    Dim sFN As String
    Dim mht As Boolean
    Dim iCounter As Integer
    
    Set WB = ThisWorkbook
    Set WS = WB.Sheets("Consumer e-library")
    iCounter = 0
    
    If Len(Dir(WB.Path & "\exported\", vbDirectory)) = 0 Then
        MkDir (WB.Path & "\exported\")
    End If
    
    For Each cll In WS.Columns(1).Cells
        sUrl = Trim(cll.Value)
        If Left(sUrl, 4) = "http" Then
            On Error Resume Next
            ext = Right(sUrl, Len(sUrl) - InStrRev(sUrl, "."))
            sFN = Right(sUrl, Len(sUrl) - InStrRev(sUrl, "/"))
            Select Case ext
                Case "htm", "html"
                    sFN = WB.Path & "\exported\" & sFN & ".mht"
                    mht = True
                Case "pdf", "jpg", "jpeg", "doc", "gif", "xls"
                    sFN = WB.Path & "\exported\" & sFN
                    mht = False
            End Select
            Select Case UrlToFile(sUrl, sFN, mht)
                Case 1
                    Log "exported: " & sFN
                    iCounter = iCounter + 1
                Case -1
                    Log "#### file/page not exists: " & sUrl
            End Select
            
            On Error GoTo 0
        End If
        If Err.Number <> 0 Then
            Log Err.Description
        End If
    Next cll
    Log "files exported: " & iCounter
    MsgBox "files exported: " & iCounter
End Sub

' save to file from url
Function UrlToFile(url As String, fn As String, asMht As Boolean) As Integer
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", url, False
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        If asMht Then
            With CreateObject("CDO.Message")
                .MimeFormatted = True
                .CreateMHTMLBody url, 0, "", ""
                .GetStream.SaveToFile fn, 2    '2 for overwrite
            End With
        Else
            With CreateObject("ADODB.Stream")
                .Open
                .Type = 1
                .write WinHttpReq.responseBody
                .SaveToFile fn, 2    '2 for overwrite
                .Close
            End With
        End If
        UrlToFile = 1   'exported successfully
    Else
        UrlToFile = -1  'page/file not exists
    End If
    
End Function

Public Function Log(str As String) As Boolean
    Open ThisWorkbook.Path & "\log.txt" For Append As #1
    Print #1, Now & ":" & vbTab & str
    Close #1
End Function


Sub test()
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", "http://www.foodsmart.govt.nz/elibrary/consumer/recall-soy-protein-products-4-11-2014.htm", False
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.write WinHttpReq.responseBody
        oStream.SaveToFile "C:\Users\Tao\Downloads\111.jpg", 2    '2 for overwrite
        oStream.Close
    End If
End Sub
