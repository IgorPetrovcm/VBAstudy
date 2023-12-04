VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3816
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5160
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub AddElementsToList(elements_arr() As elements)
    Dim index As Integer
    For index = LBound(elements_arr) To UBound(elements_arr)
        ListBox1.AddItem (elements_arr(index).name)
    Next
End Sub


Private Sub UserForm_Click()

End Sub
