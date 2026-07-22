Attribute VB_Name = "Examples"
Option Explicit
' Author: M.Kausar '
' Date: 07-Jul-2026 '
' You can modify it According to your need '

' Requirements: Reference to "Microsoft Scripting Runtime" '
'    goto Tools > Reference: Search for "Microsoft Scripting Runtime" and select '


Private Sub PrintExamples()
    ' Note: Do not add any object (Dictionary or Collection) to Array '
    ' this will break the code '
    ' eg -> Dim Arr(0 to 1): Set Arr(0) = Dictionary or Collection '
    Dim c1 As New Collection, c2 As New Collection
    Dim d1 As New Dictionary, d2 As New Dictionary
    
    Dim i&: For i = 1 To 4
        c1.Add i
        c2.Add Chr$(i + 64)
        d1.Add Chr$(i + 64), "Chr code of " & Chr$(i + 64) & " is -> " & i + 64
        d2.Add "Key" & i, "Item" & i
    Next
    
    c1.Add c2 'collection inside collection'
    
    d1.Add "Dict", d2 'dictionary inside dictionary'
    d1.Add "Coll", c2 'collection inside dictionary'
    
    c1.Add d1 'dictionary inside collection'
    
    Debug.Print PrintDictOrCollOrArr(c1) ' -> flat print '
    Debug.Print PrintDictOrCollOrArr(c1, 4) ' -> pretty print '
    Debug.Print PrintDictOrCollOrArr(c1, 4, 1) ' -> pretty print with indent '
End Sub

Private Sub CompareExamples()
    Dim c1 As New Collection, c2 As New Collection, c3 As New Collection, c4 As New Collection
    Dim d1 As New Dictionary, d2 As New Dictionary
    
    Dim i&: For i = 1 To 5
        c1.Add i
        c2.Add i
        c3.Add i * 5
        c4.Add i * 5
    Next i
    
    c4.Add 30 'add new value to collection'
    
    c1.Add c3 'collection inside another collection'
    c2.Add c4 'collection inside another collection'
    
    'below will print true. because the top level collections are equal'
    Debug.Print CompareDictOrCollOrArr(c1, c2) 'shallow compare'
    
    'below will print false. because, nested collections are not equals'
    Debug.Print CompareDictOrCollOrArr(c1, c2, True) 'deep compare'
    
    d1.Add "Coll", c1
    d2.Add "Coll", c2
    
    Debug.Print CompareDictOrCollOrArr(d1, d2) 'shallow compare'
    Debug.Print CompareDictOrCollOrArr(d1, d2, True) 'deep compare'
End Sub


Private Sub CloneExamples()
    Dim c1 As New Collection, c2 As New Collection, c3 As New Collection, c4 As New Collection
    Dim d1 As New Dictionary, d2 As New Dictionary, d3 As New Dictionary, d4 As New Dictionary
    
    Dim i&: For i = 1 To 5
        c1.Add i
        c3.Add i
    Next i
    
    Set c2 = c1 '"c1" and "c2" both point to the same object in memory, modifying one will _
                affect another'
    c2.Add "new item" ' also affect "c1" '
    Debug.Print PrintDictOrCollOrArr(c1) ' -> [1, 2, 3, 4, 5, "new item"] '
    
    Set c4 = CloneDictOrCollOrArr(c3) '"c3" and "c4" both point to the different object in memory, _
                                        modifying one won't affect another at top level'
    c4.Add "new item" ' won't affect "c3" '
    Debug.Print PrintDictOrCollOrArr(c3) ' -> [1, 2, 3, 4, 5] '
    
    For i = 1 To 2
        d1.Add Chr$(i + 64), i + 64
        d3.Add Chr$(i + 64), i + 64
    Next
    
    d1.Add "Coll", c1
    Set d2 = CloneDictOrCollOrArr(d1) ' clone top level only, modifying nested collections will affect _
                                        other nested collections '
    d2.Add "New Key", "New Item" 'won't affect "d1"'
    d2("Coll").Add "New Value" 'will affect nested collection in "d1", because not deeply cloned'

    Debug.Print PrintDictOrCollOrArr(d1, 4) ' nested collection inside d1 is changed '
    
    d3.Add "Coll", c1
    Set d4 = CloneDictOrCollOrArr(d3, True) ' deeply cloned, modifying nested collections won't affect _
                                        other nested collections '
    d4.Add "New Key", "New Item" 'won't affect "d3"'
    d4("Coll").Add "New Value" 'won't affect nested collection in "d3", because deeply cloned'
    
    Debug.Print PrintDictOrCollOrArr(d3, 4) ' nested collection inside d3 is not changed '
End Sub
















