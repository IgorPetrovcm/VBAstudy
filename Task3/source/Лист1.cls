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
Public range_start As Range
Public range_end As Range

Private Sub ListBox1_Click()
    Cells(5, 5).Value = ListBox1.Selected
End Sub

Private Sub Worksheet_Activate()
    Set range_start = Cells(1, 1)
    Set range_end = range_start.End(xlDown)
    
    'Dim index As Integer
    'For index = range_start.Row To range_end.Row
    '    ListBox1.AddItem (Cells(range_start.Row + index, range_start.Column).Value)
        
    'Next
    ListBox1.RowSource = "Sheet1!" + Replace(range_start.Address, "$", "") + ":" + Replace(range_end.Address, "$", "")
    'MsgBox ("!Sheet1" + range_start.Address + ";;;" + range_end.Address)
    Dim str As String
    str = range_start.Address
    str = Replace(str, "$", "")
    MsgBox (str)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
