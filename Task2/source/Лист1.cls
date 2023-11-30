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
    Dim size As Integer
    size = 0
    Dim unit As range
    Set unit = Cells(2, 1)
    Do While unit.Value <> ""
        size = size + 1
        Set unit = unit.Offset(1, 0)
    Loop
    MsgBox (size)
End Sub

