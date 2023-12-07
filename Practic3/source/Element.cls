VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Public price As Double
Public address As Range
Public sheet As Worksheet

Public Function Constructor(name As String, price As Double, address As Range, sheet As Worksheet) As element
    Dim element As element
    Set element = New element
    element.name = name
    element.price = price
    element.address = address
    element.sheet = sheet
    
    Set Constructor = element
End Function
