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


Private Sub ListBox1_AfterUpdate()
    Dim name_element As String
    name_element = Me.ListBox1
    
    Dim index As Integer
    For index = LBound(elements) To UBound(elements)
        If elements(index).name = name_element Then
            TextBox1.Value = elements(index).price
        End If
    Next
End Sub

Private Sub ListBox1_Click()


End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    UserForm2.Show
    UserForm2.TextBox1.Value = ListBox1.List(ListBox1.ListIndex)
End Sub

Private Sub UserForm_Activate()
    elements = Module1.GetArrayElements()

    If Me.ListBox1.ListCount = 0 Then
        Dim index As Integer
        For index = LBound(elements) To UBound(elements)
            ListBox1.AddItem (elements(index).name)
        Next
    End If
    
End Sub

Private Sub UserForm_Click()

End Sub
