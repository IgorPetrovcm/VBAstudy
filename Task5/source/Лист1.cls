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
Public start_ As Range
Public end_ As Range
Private elements() As elements

Private Sub Worksheet_Activate()
    Set start_ = Cells(1, 1)
    Set end_ = start_.End(xlDown)
    ReDim elements(end_.Row - (start_.Row - 1))
    
    Dim index As Integer
    For index = start_.Row - 1 To end_.Row
        Dim element As New elements
        Dim range_element As Range
        Set range_element = Cells(start_.Row + index, start_.Column)
        
        element.name = range_element.Value
        element.price = range_element.Offset(0, 1).Value
        Set element.address = range_element
        Set element.sheet_on_address = ActiveSheet
                
        Set elements(index) = element
    Next
End Sub


