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

Public Sub Constructor(name As String, price As Double, address As Range, sheet As Worksheet)
    Me.name = name
    Me.price = price
    Me.address = address
    Me.sheet = sheet
End Sub
