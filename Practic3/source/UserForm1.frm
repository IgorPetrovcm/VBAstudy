VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6444
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private elements() As element

Private Sub ListBox1_Click()
    
End Sub

Private Sub UserForm_Activate()
    elements = Module1.GetArrayElements()

    
    Dim index As Integer
    For index = LBound(elements) To UBound(elements)
        ListBox1.AddItem (elements(index).name)
    Next
    
End Sub

Private Sub UserForm_Click()

End Sub
