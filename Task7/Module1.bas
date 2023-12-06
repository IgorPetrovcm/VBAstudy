Attribute VB_Name = "Module1"
Option Explicit

Public Sub Ask()
    Dim element As element
    Dim arr(5) As element
    Dim index As Integer
    
    For index = LBound(arr) To UBound(arr)
        Set element = New element
        element.price = index
        Set arr(index) = element
    Next
End Sub
