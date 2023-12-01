Attribute VB_Name = "Module1"
Option Explicit
Public range_values As range
Sub test_EndRange()
    Dim endrange As range
    Dim startrange As range
    Set startrange = Cells(2, 1)
    Set endrange = startrange.End(xlDown)
    Set range_values = range(startrange, endrange)
    
    MsgBox range_values.Address
End Sub
