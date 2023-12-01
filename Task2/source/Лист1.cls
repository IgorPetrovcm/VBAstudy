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
    Dim unit As range
    Dim index As Integer
    index = 0
    Dim arr() As String
    
    Set unit = Cells(2, 1)
    Do While unit.Value <> ""
        ReDim arr(1)
        arr(index) = unit.Value
        index = index + 1
    Loop
End Sub

