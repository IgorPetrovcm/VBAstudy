Attribute VB_Name = "Module1"
Option Explicit

Public Function GetArray() As Integer()
    Dim elements(3) As Integer
    elements(0) = 1
    elements(1) = 2
    elements(2) = 3
    GetArray = elements
End Function

Sub Module()
    Dim arr() As Integer
    arr = GetArray()
    
End Sub
