Attribute VB_Name = "ArrListExamples"
Option Explicit

' fixed array '
Private Sub FixedArray()
    Dim Arr(1 To 5) As Variant
    
    Dim i&: For i = 1 To 5 ' adding more than 5 items will raise subscript out of range error '
        Arr(i) = i
    Next
    
    Debug.Print Join(Arr, vbLf) 'print to immediate window Ctrl + G'
End Sub

' dynamic array '
Private Sub DynamicArray()
    Dim Arr As New ArrList
    
    Dim i&: For i = 1 To 20
        Arr.Append i
    Next
    
    Debug.Print Join(Arr.GetArr(), vbLf) 'print to immediate window Ctrl + G'
    
    Arr.Clear ' clear array '
    
    Debug.Print Join(Arr.GetArr(), vbLf) 'nothing will print to immediate window'
End Sub

