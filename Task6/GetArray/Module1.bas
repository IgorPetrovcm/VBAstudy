Attribute VB_Name = "Module1"
Option Explicit

Public Function GetElementsArray() As element()
    Dim this As Worksheet
    Set this = ThisWorkbook.Worksheets(1)
    Dim start_ As Range
    Set start_ = this.Cells(1, 1)
    Dim end_ As Range
    Set end_ = start_.End(xlDown)
    
    Dim elements() As element
    ReDim elements(end_.Row - (start_.Row - 1))
    
    Dim index As Integer
    For index = LBound(elements) To UBound(elements)
        Dim element As New element
        Set elements(index) = element
        Dim range_element As Range
        Set range_element = this.Cells(start_.Row + index, start_.Row)
        
        element.name = range_element.Value
        Set element.address = range_element
        element.price = range_element.Offset(0, 1)
        Set element.sheet = this
    Next
    
    GetElementsArray = elements
    
End Function


