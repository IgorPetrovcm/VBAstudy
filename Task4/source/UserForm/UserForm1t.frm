VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4188
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5280
   OleObjectBlob   =   "UserForm1t.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListBox1_Click()
    
End Sub

Private Sub UserForm_Activate()
    Dim start_range As Range
    Dim end_range As Range
    Set start_range = Worksheets(1).Cells(1, 1)
    Set end_range = start_range.End(xlDown)
    
    ListBox1.RowSource = "Sheet1!" & Replace(start_range.Address, "$", "") & ":" & Replace(end_range.Address, "$", "")
End Sub

Private Sub UserForm_Click()

End Sub
