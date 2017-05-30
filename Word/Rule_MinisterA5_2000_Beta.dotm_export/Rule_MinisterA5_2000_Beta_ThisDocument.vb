VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    Dim cc As ContentControl
    Dim doc As Document
    Set doc = ActiveDocument
    Dim strUnit As String
    Dim strNumber As String
    
    If ContentControl.Title = "EffectiveDate" Then
        For Each cc In doc.SelectContentControlsByTitle("OffsetDate")
            strUnit = Left(cc.Tag, 1)
            strNumber = Right(cc.Tag, Len(cc.Tag) - 1)
            If strUnit = "m" Or strUnit = "d" Or strUnit = "y" Then
                If IsNumeric(strNumber) Then
                    If Int(strNumber) = strNumber Then
                        If strUnit = "y" Then
                            strUnit = "yyyy"
                        End If
                        cc.Range.Text = Format(DateAdd(strUnit, strNumber, ContentControl.Range.Text), DateFormat)
                    End If
                End If
            End If
        Next cc
    End If
End Sub



