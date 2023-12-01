Attribute VB_Name = "Module1"
Option Explicit
Public range_values As range
Sub test_EndRange()
    Dim endrange As range
    Dim startrange As range
    Set startrange = Cells(2, 1)
    Set endrange = startrange.End(xlDown)
    Set range_values = range(startrange, endrange)
    
    Dim size_arr As Integer
    Dim index As Integer
    
    size_arr = range_values.Rows.Count
    Dim arr() As String
    ReDim arr(range_values.Rows.Count)
    For index = 0 To UBound(arr)
        arr(index) = Cells(startrange.Row + index, startrange.Column).Value
    Next
    
End Sub
