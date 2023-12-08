Attribute VB_Name = "Module2"
Option Explicit

Function GetCounting(ByVal name_element As String, elements() As element) As element
    Dim index As Integer
    For index = LBound(elements) To UBound(elements)
        If elements(index) = name_element Then
            GetCounting = elements(index)
        End If
    Next
    GetCounting (elemetns(0))
End Function
