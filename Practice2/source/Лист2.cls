VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Public start_cell As Range
Public end_cell As Range

Private Sub Worksheet_Activate()
    Cells(2, 2).Value = "true"
    
    Set start_cell = Cells(1, 1)
    Set end_cell = start_cell.End(xlDown)
    
    Dim index As Integer
     
    For index = 0 To end_cell.Row
        ListBox1.AddItem (Cells(start_cell.Row + index, start_cell.Column))
        
    Next
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub

