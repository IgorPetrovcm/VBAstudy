VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Dim arr(0 To 19) As String
    Dim unit As range
    Set unit = Cells(2, 1)
    For i as Integer = 1 To 19
    arr [Index] = unit.Value
    Set unit = unit.Offset(1, 0)
    Next
    MsgBox (size)
End Sub

