# 📦 VBA JSON & Data Tools  
**Deep Compare & Deep Clone for Dictionary, Collection, Array (up to 12D)**  

---

## 🚀 Features
- **JSON Conversion** – Convert `Dictionary`, `Collection`, or `Array` (up to 12D) into JSON.  
- **Deep Compare** – Recursively compare nested structures.  
- **Deep Clone** – Create independent copies of complex objects.  
- **Zero Dependencies** – Works in Excel, Access, Word, PowerPoint VBA 2010+.  

---

## 📥 Installation
1. Download `DictCollArrMaster.bas`  
2. In VBE: `File > Import File > DictCollArrMaster.bas`  
3. Add reference: `Tools > References > Microsoft Scripting Runtime`  

---

1. **PrintDictCollArr** - Convert to JSON
Convert `Scripting.Dictionary`, `Collection`, or any `Variant Array` up to 12 dimensions into a JSON string.  
Handles unlimited nesting and mixed types.

2. **CompareDictCollArr** - Deep Compare
Compare two Dictionaries, Collections, or Arrays.  
Choose shallow or deep recursive comparison.

3. **CloneDictCollArr** - Deep Clone  
Create a shallow or true deep copy of any Dictionary, Collection, or Array.  
Nested objects are cloned too when `DeepClone = True`.

---

**API Reference**

**1. PrintDictCollArr**
```vba
Public Function PrintDictCollArr( _
    ByVal Iterable As Variant, _
    Optional ByVal TabSize As Byte = 0, _
    Optional ByVal Indent As Byte = 0 _
) As String
```
Convert any iterable to JSON.

*Parameters:*
Param Type Default Description
`Iterable` Variant Required Dictionary, Collection, or Array 1D to 12D
`TabSize` Byte 0 `0` = Minified JSON. `>0` = Pretty print with that many spaces per indent level
`Indent` Byte 0 Number of indent levels to add to the start of each line. Uses `TabSize` spaces
*Examples:*
```vba
Dim arr As Variant: arr = Array(1, 2, 3, 4, 5)


' Minified
Debug.Print PrintDictCollArr(arr) 
' [1,2,3,4,5]

' Pretty Print with 4 spaces
Debug.Print PrintDictCollArr(arr, TabSize:=4)
' [
'   1,
'   2,
'   3,
'   4,
'   5
' ]

' Pretty Print with 4 spaces + 2 levels of indent
Debug.Print PrintDictCollArr(arr, TabSize:=4, Indent:=2)
'     [
'       1,
'       2,
'       3,
'       4,
'       5
'     ]
```
*2. CompareDictCollArr*
```vba
Public Function CompareDictCollArr( _
    ByVal Iterable1 As Variant, _
    ByVal Iterable2 As Variant, _
    Optional ByVal DeepSearch As Boolean = False _
) As Boolean
```
*Parameters:*
Param Type Default Description
`Iterable1` Variant Required First Dictionary, Collection, or Array
`Iterable2` Variant Required Second Dictionary, Collection, or Array
`DeepSearch` Boolean False `False` = Shallow compare. `True` = Recursively compare all nested items
*Returns:* `True` if both structures and all values match.
```vba
If CompareDictCollArr(dict1, dict2, DeepSearch:=True) Then
    Debug.Print "Identical"
End If
```
*3. CloneDictCollArr*
```vba
Public Function CloneDictCollArr( _
    ByVal Iterable As Variant, _
    Optional ByVal DeepClone As Boolean = False _
) As Object
```
*Parameters:*
Param Type Default Description
`Iterable` Variant Required Dictionary, Collection, or Array to clone
`DeepClone` Boolean False `False` = Shallow copy. `True` = Recursively clone all nested objects/arrays
*Returns:* New `Object`. For arrays it returns a `Variant()` that you `Set` is not needed.
```vba
Dim copy As Object
Set copy = CloneDictCollArr(originalDict, DeepClone:=True)

copy("key")("nested") = "changed" ' original is not affected
```

*Technical Specs*
Feature Detail
**Max Array Dimensions** 12
**Max Nesting** Unlimited, limited by VBA call stack
**Supported Types** Dictionary, Collection, Array, String, Number, Boolean, Date, Null, Empty
**Date Format in JSON** ISO 8601: `yyyy-mm-ddThh:nn:ss`
**Dependencies** Microsoft Scripting Runtime
---

*Use Cases*
- Export Excel data to JSON for APIs
- Save/Load user settings to a .json file
- Unit testing VBA code with `CompareDictCollArr`
- Create templates and working copies with `CloneDictCollArr`
- Debug complex nested data with pretty-printed JSON

---

*License*
MIT License. Free for personal and commercial use.

If this module helped you, please give it a ⭐

*Author*: M.Kausar
