VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public sQuestion As String
Public ActionNo As Integer
Public sAnswer As String
Public sTip As String
Public NeedAnswer As Boolean
Public YesNode As String    'name of the node when 'Yes' selected
Public NoNode As String     'name of the node when 'No' selected
Public PreviousNode As String   'previous node to trace back
Public NextNode As String
Public YesNo As String      'store Yes/No selection, could be 'y' 'n' or blank(stands for no choice made)

Private Sub Class_Initialize()
    'populate controls in form
End Sub
