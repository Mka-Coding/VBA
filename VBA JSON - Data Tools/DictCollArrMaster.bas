Attribute VB_Name = "DictCollArrMaster"
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

' Requirements: Reference to "Microsoft Scripting Runtime" '
'    goto Tools > Reference: Search for "Microsoft Scripting Runtime" and select '

'======================================'
' PUBLIC FUNCTIONS '
'======================================'
Public Function CloneDictOrCollOrArr(ByRef Iterable As Variant, Optional DeepClone As Boolean = False) As Variant
    If VBA.TypeName(Iterable) = "Dictionary" Then
        Set CloneDictOrCollOrArr = CloneDictionary(Iterable, DeepClone)
    ElseIf VBA.TypeName(Iterable) = "Collection" Then
        Set CloneDictOrCollOrArr = CloneCollection(Iterable, DeepClone)
    ElseIf VBA.IsArray(Iterable) Then
        Let CloneDictOrCollOrArr = CloneArray(Iterable, DeepClone) ' Arrays Always DeepClone: DeepClone Can't Be False
    Else
        MsgBox "Can't Clone Object Type: '" & VBA.TypeName(Iterable) & "'" & vbLf & "Expecting 'Dictionary' or 'Collection' or 'Array'", vbCritical, "Error"
    End If
End Function
Public Function CompareDictOrCollOrArr(ByRef Iterable1 As Variant, ByRef Iterable2 As Variant _
                                , Optional DeepSearch As Boolean = False) As Boolean
    If VBA.TypeName(Iterable1) = "Dictionary" And IsIterable(Iterable2) Then
        Let CompareDictOrCollOrArr = CompareDictionaries(Iterable1, Iterable2, DeepSearch)
    ElseIf VBA.TypeName(Iterable1) = "Collection" And IsIterable(Iterable2) Then
        Let CompareDictOrCollOrArr = CompareCollections(Iterable1, Iterable2, DeepSearch)
    ElseIf VBA.IsArray(Iterable1) And IsIterable(Iterable2) Then
        Let CompareDictOrCollOrArr = CompareArrays(Iterable1, Iterable2, DeepSearch)
    Else
        If Not IsIterable(Iterable1) Then
            MsgBox "Can't Compare Object Type: '" & VBA.TypeName(Iterable1) & "'" & vbLf & "Expecting 'Dictionary' or 'Collection' or 'Array'", vbCritical, "Error"
        Else
            MsgBox "Can't Compare Object Type: '" & VBA.TypeName(Iterable2) & "'" & vbLf & "Expecting 'Dictionary' or 'Collection' or 'Array'", vbCritical, "Error"
        End If
    End If
End Function
Public Function PrintDictOrCollOrArr(ByRef Iterable As Variant, Optional ByRef TabSize As Byte = 0 _
                                , Optional ByRef Indent As Byte = 0) As String
    If VBA.TypeName(Iterable) = "Dictionary" Then
        Let PrintDictOrCollOrArr = PrintDict(Iterable, TabSize:=TabSize, Indent:=Indent)
    ElseIf VBA.TypeName(Iterable) = "Collection" Then
        Let PrintDictOrCollOrArr = PrintColl(Iterable, TabSize:=TabSize, Indent:=Indent)
    ElseIf VBA.IsArray(Iterable) Then
        Let PrintDictOrCollOrArr = PrintArr(Iterable, TabSize:=TabSize, Indent:=Indent)
    Else
        MsgBox "Can't Print Object Type: '" & VBA.TypeName(Iterable) & "'" & vbLf & "Expecting 'Dictionary' or 'Collection' or 'Array'", vbCritical, "Error"
    End If
End Function
'======================================'
' PUBLIC FUNCTIONS '
'======================================'

'======================================'
' CLONE ARRAY '
'======================================'
Private Function CloneArray(ByVal Arr As Variant, Optional DeepClone As Boolean = False) As Variant
    Let CloneArray = Arr
End Function
'======================================'
' CLONE ARRAY '
'======================================'

'======================================'
' CLONE DICTIONARY OR ARRAY '
'======================================'
Private Function CloneDictionary(ByVal Dict As Object, Optional DeepClone As Boolean = False) As Dictionary
    'Dim CloneDict As Object: Set CloneDict = CreateObject("Scripting.Dictionary") 'not working on mac'
    Dim CloneDict As New Dictionary ' required reference to "Microsoft Scripting Runtime" '
    Dim Key: For Each Key In Dict
        If VBA.TypeName(Dict(Key)) = "Dictionary" And DeepClone Then
            CloneDict.Add Key, CloneDictionary(Dict(Key), DeepClone)
        ElseIf VBA.TypeName(Dict(Key)) = "Collection" And DeepClone Then
            CloneDict.Add Key, CloneCollection(Dict(Key), DeepClone)
        Else
            CloneDict.Add Key, Dict(Key)
        End If
    Next
    Set CloneDictionary = CloneDict
End Function
'======================================'
' CLONE DICTIONARY OR ARRAY '
'======================================'

'======================================'
' CLONE COLLECTION OR ARRAY '
'======================================'
Private Function CloneCollection(ByVal Coll As Collection, Optional DeepClone As Boolean = False) As Collection
    Dim CloneColl As New Collection
    Dim Itm: For Each Itm In Coll
        If VBA.TypeName(Itm) = "Collection" And DeepClone Then
            CloneColl.Add CloneCollection(Itm, DeepClone)
        ElseIf VBA.TypeName(Itm) = "Dictionary" And DeepClone Then
            CloneColl.Add CloneDictionary(Itm, DeepClone)
        Else
            CloneColl.Add Itm
        End If
    Next
    Set CloneCollection = CloneColl
End Function
'======================================'
' CLONE COLLECTION OR ARRAY '
'======================================'



'======================================'
' COMPARE DICTIONARY '
'======================================'
Private Function CompareDictionaries(ByVal Dict1 As Object, ByVal Dict2 As Object, Optional DeepSearch As Boolean = False) As Boolean
    CompareDictionaries = (VBA.TypeName(Dict1) = VBA.TypeName(Dict2) And Dict1.count = Dict2.count)
    If Not CompareDictionaries Then Exit Function
    
    Dim Key: For Each Key In Dict1.Keys
        If VBA.TypeName(Dict1(Key)) <> VBA.TypeName(Dict2(Key)) Then CompareDictionaries = False: Exit For
        If DeepSearch Then
            If VBA.TypeName(Dict1(Key)) = "Dictionary" Then
                CompareDictionaries = CompareDictionaries(Dict1(Key), Dict2(Key), DeepSearch)
                If Not CompareDictionaries Then Exit For
                GoTo Continue
            ElseIf VBA.TypeName(Dict1(Key)) = "Collection" Then
                CompareDictionaries = CompareCollections(Dict1(Key), Dict2(Key), DeepSearch)
                If Not CompareDictionaries Then Exit For
                GoTo Continue
            ElseIf VBA.IsArray(Dict1(Key)) Then
                CompareDictionaries = CompareArrays(Dict1(Key), Dict2(Key), DeepSearch)
                If Not CompareDictionaries Then Exit For
                GoTo Continue
            End If
        Else
            If VBA.TypeName(Dict1(Key)) = "Dictionary" Or VBA.TypeName(Dict1(Key)) = "Collection" Or VBA.IsArray(Dict1(Key)) Then GoTo Continue
        End If
        If Dict1(Key) <> Dict2(Key) Then
            CompareDictionaries = False
            Exit For
        End If
Continue:
    Next
End Function
'======================================'
' COMPARE DICTIONARY '
'======================================'

'======================================'
' COMPARE COLLECTION '
'======================================'
Private Function CompareCollections(ByVal Coll1 As Collection, ByVal Coll2 As Collection, Optional DeepSearch As Boolean = False) As Boolean
    CompareCollections = (VBA.TypeName(Coll1) = VBA.TypeName(Coll2) And Coll1.count = Coll2.count)
    If Not CompareCollections Then Exit Function
    
    Dim Idx&: For Idx = 1 To Coll1.count
        If VBA.TypeName(Coll1(Idx)) <> VBA.TypeName(Coll2(Idx)) Then CompareCollections = False: Exit For
        If DeepSearch Then
            If VBA.TypeName(Coll1(Idx)) = "Collection" Then
                CompareCollections = CompareCollections(Coll1(Idx), Coll2(Idx), DeepSearch)
                If Not CompareCollections Then Exit For
                GoTo Continue
            ElseIf VBA.TypeName(Coll1(Idx)) = "Dictionary" Then
                CompareCollections = CompareDictionaries(Coll1(Idx), Coll2(Idx), DeepSearch)
                If Not CompareCollections Then Exit For
                GoTo Continue
            ElseIf VBA.IsArray(Coll1(Idx)) Then
                CompareCollections = CompareArrays(Coll1(Idx), Coll2(Idx), DeepSearch)
                If Not CompareCollections Then Exit For
                GoTo Continue
            End If
        Else
            If VBA.TypeName(Coll1(Idx)) = "Collection" Or VBA.TypeName(Coll1(Idx)) = "Dictionary" Or VBA.IsArray(Coll1(Idx)) Then GoTo Continue
        End If
        If Coll1(Idx) <> Coll2(Idx) Then
            CompareCollections = False
            Exit For
        End If
Continue:
    Next
End Function
'======================================'
' COMPARE COLLECTION '
'======================================'

'======================================'
' COMPARE ARRAY '
'======================================'
Private Function CompareArrays(ByVal Arr1 As Variant, ByVal Arr2 As Variant, Optional DeepSearch As Boolean = False) As Boolean
    CompareArrays = (VBA.TypeName(Arr1) = VBA.TypeName(Arr2))
    If Not CompareArrays Then Exit Function
    CompareArrays = (LBound(Arr1) = LBound(Arr2) And UBound(Arr1) = UBound(Arr2) And GetDimensions(Arr1) = GetDimensions(Arr2))
    If Not CompareArrays Then Exit Function
    Dim Dimensions(): ReDim Dimensions(1 To GetDimensions(Arr1))
    CompareArrays = CompareRecursive(Arr1, Arr2, Dimensions, DeepSearch)
End Function
Private Function CompareRecursive(Arr1 As Variant, Arr2 As Variant, Dimensions As Variant _
                            , Optional DeepSearch As Boolean = False, Optional ByVal Depth As Byte = 1) As Boolean
    Dim i&: For i = LBound(Arr1, Depth) To UBound(Arr2, Depth)
        Dimensions(Depth) = i
        
        If Depth = UBound(Dimensions) Then
            Dim Item1: Item1 = GetItem(Arr1, Dimensions)
            Dim Item2: Item2 = GetItem(Arr2, Dimensions)
            CompareRecursive = VBA.TypeName(Item1) = VBA.TypeName(Item2)
            If Not CompareRecursive Then Exit Function
            
            If VBA.IsArray(Item1) Then
                If DeepSearch Then
                    ' go deeper '
                    CompareRecursive = CompareArrays(Item1, Item2, DeepSearch)
                    If Not CompareRecursive Then Exit Function
                End If
            Else
                ' reach the last dimension. this is real element '
                CompareRecursive = (Item1 = Item2)
                If Not CompareRecursive Then Exit Function
            End If
        Else
            ' goto next dimension '
            CompareRecursive = CompareRecursive(Arr1, Arr2, Dimensions, DeepSearch, Depth + 1)
            If Not CompareRecursive Then Exit Function
        End If
    Next
End Function
'======================================'
' COMPARE ARRAY '
'======================================'


'======================================'
' PRINT DICTIONARY '
'======================================'
Private Function PrintDict(ByVal Dict As Object, Optional ByVal DictKey As String = "", _
                            Optional ByVal TabSize As Byte = 0, Optional ByVal Indent As Byte = 0) As String
    Dim Key As Variant, Index As Long
    Dim ArrSize&: ArrSize = Dict.count + 2
    Dim ArrIdx&, ArrList$(): ReDim ArrList$(0 To ArrSize)
    
    If DictKey = "" Then
        AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + "{" + IIf(TabSize = 0, " ", "")
    Else
        AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + ParseKey(DictKey) _
            + ": {" + IIf(TabSize = 0, " ", "")
    End If
    For Each Key In Dict.Keys
        If VBA.TypeName(Dict(Key)) = "Dictionary" Then
            AppendArr ArrList, ArrIdx, ArrSize, PrintDict(Dict(Key), Key, TabSize, Indent + 1) _
                + IIf(Index = Dict.count - 1, "", IIf(TabSize = 0, ", ", ","))
        ElseIf VBA.TypeName(Dict(Key)) = "Collection" Then
            AppendArr ArrList, ArrIdx, ArrSize, PrintColl(Dict(Key), Key, TabSize, Indent + 1) _
                + IIf(Index = Dict.count - 1, "", IIf(TabSize = 0, ", ", ","))
        ElseIf VBA.IsArray(Dict(Key)) Then
            AppendArr ArrList, ArrIdx, ArrSize, PrintArr(Dict(Key), Key, TabSize, Indent + 1) _
                + IIf(Index = Dict.count - 1, "", IIf(TabSize = 0, ", ", ","))
        Else
            AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize + TabSize) + ParseKey(Key) _
                + ": " + ParseValue(Dict(Key)) + IIf(Index = Dict.count - 1, "", IIf(TabSize = 0, ", ", ","))
        End If
        Let Index = Index + 1
    Next Key
    AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + IIf(TabSize = 0, " ", "") + "}"
    
    ReDim Preserve ArrList$(0 To ArrIdx - 1)
    Let PrintDict = VBA.Join(ArrList, IIf(TabSize = 0, "", vbLf))
End Function
'======================================'
' PRINT DICTIONARY '
'======================================'

'======================================'
' PRINT COLLECTION '
'======================================'
Private Function PrintColl(ByVal Coll As Collection, Optional ByVal DictKey As String = "", _
                            Optional ByVal TabSize As Byte = 0, Optional ByVal Indent As Byte = 0) As String
    Dim Itm As Variant, Index As Long
    Dim ArrSize&: ArrSize = Coll.count + 2
    Dim ArrIdx&, ArrList$(): ReDim ArrList$(0 To ArrSize)

    If DictKey = "" Then
        AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + "["
    Else
        AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + ParseKey(DictKey) + ": ["
    End If
    For Each Itm In Coll
        If VBA.TypeName(Itm) = "Collection" Then
            AppendArr ArrList, ArrIdx, ArrSize, PrintColl(Itm, TabSize:=TabSize, Indent:=Indent + 1) _
                + IIf(Index = Coll.count - 1, "", IIf(TabSize = 0, ", ", ","))
        ElseIf VBA.TypeName(Itm) = "Dictionary" Then
            AppendArr ArrList, ArrIdx, ArrSize, PrintDict(Itm, TabSize:=TabSize, Indent:=Indent + 1) _
                + IIf(Index = Coll.count - 1, "", IIf(TabSize = 0, ", ", ","))
        ElseIf VBA.IsArray(Itm) Then
            AppendArr ArrList, ArrIdx, ArrSize, PrintArr(Itm, TabSize:=TabSize, Indent:=Indent + 1) _
                + IIf(Index = Coll.count - 1, "", IIf(TabSize = 0, ", ", ","))
        Else
            AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize + TabSize) + ParseValue(Itm) _
                + IIf(Index = Coll.count - 1, "", IIf(TabSize = 0, ", ", ","))
        End If
        Let Index = Index + 1
    Next Itm
    AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + "]"
    
    ReDim Preserve ArrList$(0 To ArrIdx - 1)
    Let PrintColl = VBA.Join(ArrList, IIf(TabSize = 0, "", vbLf))
End Function
'======================================'
' PRINT COLLECTION '
'======================================'

'======================================'
' PRINT ARRAY '
'======================================'
Private Function PrintArr(Arr As Variant, Optional ByVal DictKey As String = "", Optional ByVal TabSize As Byte = 0 _
                        , Optional ByVal Indent As Byte = 0, Optional IsLast As Boolean = True) As String
    
    Dim ArrSize&: ArrSize = UBound(Arr) + 2
    Dim ArrIdx&, ArrList$(): ReDim ArrList$(0 To ArrSize)
    
    If DictKey = "" Then
        AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + "["
    Else
        AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + ParseKey(DictKey) + ": ["
    End If
    
    Dim Dimensions(): ReDim Dimensions(1 To GetDimensions(Arr))
    AppendArr ArrList, ArrIdx, ArrSize, PrintRecursive(Arr, Dimensions, TabSize, Indent)
    AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize) + "]" + IIf(IsLast, "", ",")
    
    ReDim Preserve ArrList$(0 To ArrIdx - 1)
    PrintArr = VBA.Join(ArrList, IIf(TabSize = 0, "", vbLf))
End Function

Private Function PrintRecursive(Arr As Variant, Dimensions As Variant, Optional TabSize As Byte = 0 _
                                , Optional Indent As Byte = 0, Optional Depth As Byte = 1) As String

    Dim ArrSize&: ArrSize = UBound(Arr) + 2
    Dim ArrIdx&, ArrList$(): ReDim ArrList$(0 To ArrSize)
    
    Dim Idx&: For Idx = LBound(Arr, Depth) To UBound(Arr, Depth)
        Dimensions(Depth) = Idx
        
        If Depth = UBound(Dimensions) Then
            Dim Item: Item = GetItem(Arr, Dimensions)
            If VBA.IsArray(Item) Then
                ' go deeper '
                AppendArr ArrList, ArrIdx, ArrSize, PrintArr(Item, TabSize:=TabSize, Indent:=Indent _
                    + 1, IsLast:=Idx = UBound(Arr, Depth))
            ' can't get object (Dictionary or Collection) from inside array '
            ' we have to use "Set" keyword if the current item is object. eg -> Set Item = GetItem(Arr, Dimensions)'
'            ElseIf VBA.TypeName(Item) = "Dictionary" Then
'                AppendArr ArrList, ArrIdx, ArrSize, PrintDict(Item, TabSize:=TabSize, Indent:=Indent + 1)
'            ElseIf VBA.TypeName(Item) = "Collection" Then
'                AppendArr ArrList, ArrIdx, ArrSize, PrintColl(Item, TabSize:=TabSize, Indent:=Indent + 1)
            Else
                ' reach the last dimension. this is real element '
                AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize + TabSize) & ParseValue(Item) _
                    & IIf(Idx = UBound(Arr, Depth), "", IIf(TabSize = 0, ", ", ","))
            End If
        Else
            ' goto next dimension '
            AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize + TabSize) + "["
            AppendArr ArrList, ArrIdx, ArrSize, PrintRecursive(Arr, Dimensions, TabSize, Indent + 1, Depth + 1)
            AppendArr ArrList, ArrIdx, ArrSize, VBA.Space$(Indent * TabSize + TabSize) + "]" _
                & IIf(Idx = UBound(Arr, Depth), "", ",")
        End If
    Next
    
    ReDim Preserve ArrList$(0 To ArrIdx - 1)
    PrintRecursive = VBA.Join(ArrList, IIf(TabSize = 0, "", vbLf))
End Function
'======================================'
' PRINT ARRAY '
'======================================'


'======================================'
' HELPER METHODS '
'======================================'
Private Sub AppendArr(ArrList$(), ArrIdx&, ArrSize&, Text$)
    If ArrIdx > ArrSize Then
        ArrSize = ArrSize * 2 + 2
        ReDim Preserve ArrList$(0 To ArrSize)
    End If
    ArrList(ArrIdx) = Text
    ArrIdx = ArrIdx + 1
End Sub
Private Sub AppendSb(Sb$, SbIdx&, SbSize&, Text$)
    Static TextLen As Long: TextLen = Len(Text)
    If SbIdx + TextLen - 1 > SbSize Then
        If SbIdx + TextLen - 1 > SbSize * 2 Then
            'Sb = Sb + VBA.Space$(TextLen - Len(Mid$(Sb, SbIdx))) 'if you don't want trailing spaces'
            Sb = Sb + VBA.Space$(TextLen)
        Else
            Sb = Sb + VBA.Space$(SbSize)
        End If
        SbSize = Len(Sb)
    End If
    Mid$(Sb, SbIdx, TextLen) = Text
    SbIdx = SbIdx + TextLen
End Sub


Private Function ParseKey$(Key)
    ParseKey = ParseString(Key)
End Function
Private Function ParseValue$(Value)
    If VBA.VarType(Value) = vbString Then
        ParseValue = ParseString(Value)
    ElseIf VBA.VarType(Value) = vbBoolean Then
        If Value Then
            ParseValue = "true"
        Else
            ParseValue = "false"
        End If
    ElseIf VBA.VarType(Value) = vbNull Or VBA.VarType(Value) = vbEmpty Then
        ParseValue = "null"
    Else
        ParseValue = Value
    End If
End Function

'+---------------------------------------------------------+'
'|       Copy from timhall reposatory json converter       |'
'+---------------------------------------------------------+'
'https://github.com/VBA-tools/VBA-JSON/blob/master/JsonConverter.bas'
Private Function ParseString$(Text)
    Dim SbSize&: SbSize = Len(Text)
    Dim SbIdx&: SbIdx = 1
    Dim Sb$: Sb = VBA.Space$(SbSize)

    Dim ChIdx&: For ChIdx = 1 To SbSize
        Dim Char$: Char = VBA.Mid$(Text, ChIdx, 1)
        Dim ChCode&: ChCode = VBA.AscW(Char)

        Select Case ChCode
        Case 34
            ' " -> 34 -> \"
            Char = "\"""
        Case 92
            ' \ -> 92 -> \\
            Char = "\\"
        Case 8
            ' backspace -> 8 -> \b
            Char = "\b"
        Case 12
            ' form feed -> 12 -> \f
            Char = "\f"
        Case 10
            ' line feed -> 10 -> \n
            Char = "\n"
        Case 13
            ' carriage return -> 13 -> \r
            Char = "\r"
        Case 9
            ' tab -> 9 -> \t
            Char = "\t"
        End Select

        AppendSb Sb, SbIdx, SbSize, Char
    Next
    
    ParseString = """" & VBA.Left$(Sb, SbIdx - 1) & """"
End Function

' if below function give you error: even i handled it. '
' goto Tools > Options > Generals '
' see in "Error Trapping" tab and select "Break on Unhandled Errors" '
Private Function GetDimensions(Arr) As Byte
    
    On Error GoTo Done
    Dim Dimension As Byte: For Dimension = 1 To 60 ' VBA max 60 dimensions
        Dim Idx As Byte: Idx = LBound(Arr, Dimension)
        GetDimensions = Dimension
    Next
Done:

End Function

'+---------------------------------------------------+'
'|       upto 12 dimensions. you can modify it       |'
'+---------------------------------------------------+'
Private Function GetItem(Arr, Dimensions)
    Select Case UBound(Dimensions) - LBound(Dimensions) + 1
        Case 1: GetItem = Arr(Dimensions(1))
        Case 2: GetItem = Arr(Dimensions(1), Dimensions(2))
        Case 3: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3))
        Case 4: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4))
        Case 5: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5))
        Case 6: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6))
        Case 7: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7))
        Case 8: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8))
        Case 9: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9))
        Case 10: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10))
        Case 11: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11))
        Case 12: GetItem = Arr(Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12))
    End Select
End Function

Private Function VbAny(StrCheck As String, ParamArray StrArgs() As Variant) As Boolean
    VbAny = False
    Static Item: For Each Item In StrArgs
        If StrCheck = Item Then
            VbAny = True
            Exit For
        End If
    Next
End Function
Private Function VbAll(StrCheck As String, ParamArray StrArgs() As Variant) As Boolean
    VbAll = True
    Static Item: For Each Item In StrArgs
        If StrCheck <> Item Then
            VbAll = False
            Exit For
        End If
    Next
End Function
Private Function IsIterable(Iterable As Variant) As Boolean
    IsIterable = VbAny(VBA.TypeName(Iterable), "Dictionary", "Collection") Or VBA.IsArray(Iterable)
End Function
'======================================'
' HELPER METHODS '
'======================================'










