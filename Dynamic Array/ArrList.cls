VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'+-------------------------------------------------------------------------------------+'
'|                                                                                     |'
'|  MIT License                                                                        |'
'|                                                                                     |'
'|  Copyright (c) 2026 MKA Coding                                                      |'
'|                                                                                     |'
'|  Permission is hereby granted, free of charge, to any person obtaining a copy       |'
'|  of this software and associated documentation files (the "Software"), to deal      |'
'|  in the Software without restriction, including without limitation the rights       |'
'|  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell          |'
'|  copies of the Software, and to permit persons to whom the Software is              |'
'|  furnished to do so, subject to the following conditions:                           |'
'|                                                                                     |'
'|  The above copyright notice and this permission notice shall be included in all     |'
'|  copies or substantial portions of the Software.                                    |'
'|                                                                                     |'
'|  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR         |'
'|  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,           |'
'|  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE        |'
'|  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER             |'
'|  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,      |'
'|  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE      |'
'|  SOFTWARE.                                                                          |'
'|                                                                                     |'
'|  Author: M.Kausar                                                                   |'
'|  Date: 22-Jul-2026                                                                  |'
'|                                                                                     |'
'+-------------------------------------------------------------------------------------+'

Private ArrIndex As Long, ArrSize As Long
Private ArrList() As Variant

Public Sub Append(Item)
    If ArrIndex > ArrSize Then
        ArrSize = ArrSize * 2 + 2 ' double the size of array + 2 buffer '
        ReDim Preserve ArrList(0 To ArrSize)
    End If
    ArrList(ArrIndex) = Item
    ArrIndex = ArrIndex + 1
End Sub

Public Function GetArr() As Variant()
    If ArrIndex > 0 Then
        GetArr = ArrList
        ReDim Preserve GetArr(0 To ArrIndex - 1)
    End If
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Private Sub Class_Initialize()
    ArrIndex = 0
    ArrSize = 7
    ReDim ArrList(ArrIndex To ArrSize)
End Sub
