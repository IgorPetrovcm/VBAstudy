Attribute VB_Name = "Module1"
Option Explicit

Public Function GetArrayElements() As element()
    Dim this_sheet As Worksheet
    Set this_sheet = ThisWorkbook.Worksheets("Елементы")
    
    Dim start_ As Range
    Set start_ = this_sheet.Cells(1, 1)
    Dim end_ As Range
    Set end_ = start_.End(xlDown)
    
    Dim elements() As element
    ReDim elements(end_.Row - start_.Row)
    Dim element As element
    
    Dim index As Byte
    For index = LBound(elements) To UBound(elements)
        Dim index_range As Range
        Set index_range = this_sheet.Cells(start_.Row + index, start_.Column)
        
        Set element = New element
        element.Constructor name:=index_range.Value, price:=index_range.Offset(0, 1).Value, address:=index_range, sheet:=this_sheet
        
        Set elements(index) = element
    Next
    
    GetArrayElements = elements
End Function
