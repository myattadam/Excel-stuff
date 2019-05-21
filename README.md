# VBA tools and macro development

## General VBA

[Microsoft VBA overview](https://docs.microsoft.com/en-us/office/vba/api/overview/)

### Arrays
You can use the ``Array`` keyword for prefilling arrays, for example ``Array("A", "B", "C")`` to create a one-dimensional array, or ``Array(Array(1, 2, 3), Array(4, 5, 6))`` to create a staggered array (an array of arrays).

You can also use the ``Evaluate`` keyword to create an array, ``Evaluate("{1, 2, 3; 4, 5, 6}")`` or the shorthand alternative, ``[{1, 2, 3; 4, 5, 6}]``, using a semicolon to create a multidimensional array. Note that when creating multidimensional array with this method, the array needs be balanced.

The evaluate keyword, when used to create an array, returns an array object so the returned value can be accessed like an array, like this: ``Evaluate("{1, 2, 3}")(1) = 2`` but you can't do the same using the shorthand method as the array is not created via a function.


Multidimensional arrays built using the evaluate function can be iterated over using a `For Each ... Next` loop:
```VB
a = Evaluate("{1, 2, 3; 4, 5, 6}") ' OR [{1, 2, 3; 4, 5, 6}]

For Each i In a
    Debug.Print i
Next

' Outputs:
' a(1, 1) = 1
' a(2, 1) = 4
' a(1, 2) = 2
' a(2, 2) = 5
' a(1, 3) = 3
' a(2, 3) = 6
```

Using ``Transpose`` on an array flips the array X/Y:
```VB
a = Application.Transpose(a)

For Each i In a
    Debug.Print i
Next

' Outputs:
' a(1, 1) = 1
' a(2, 1) = 2
' a(3, 1) = 3
' a(1, 2) = 4
' a(2, 2) = 5
' a(3, 2) = 6
```
Using a ``For ... Next`` loop doesn't work the same way because of the multiple dimensions; you need a loop for each dimension of the array.


the following using the ``Array()`` function however _does not_ work, because the Array keyword creates a staggered array `(1)(1)` rather than a multidimensional array `(1,1)`:
```    
  a = Array(Array(1, 2, 3), Array(4, 5, 6))
  
  For Each i In a
      Debug.Print i
  Next
```

#### Sorting arrays
This function takes an one-dimensional or an staggered array, sorts them ascendingly, and passes back the sorted array. If you are passing a one-dimensional array, there's no need to specify the column to sort by (`byColumn`).

```basic
Function QuickSort(arr As Variant, Optional byColumn As Long = -1) As Variant
    Dim left As Variant
    Dim right As Variant
    Dim pivot As Variant
    
    Dim i As Long
    
    If Arrays.Length(arr) > 1 Then
    
        pivot = arr(UBound(arr))
        
        For i = LBound(arr) To UBound(arr) - 1
            If byColumn = -1 Then
        
                If arr(i) <= pivot Then
                    Arrays.Append left, arr(i)
                Else
                    Arrays.Append right, arr(i)
                End If
                
            Else
            
                If arr(i)(byColumn) <= pivot(byColumn) Then
                    Arrays.Append left, arr(i)
                Else
                    Arrays.Append right, arr(i)
                End If
                
            End If
        Next
        
        QuickSort left, byColumn
        QuickSort right, byColumn
    
        arr = Empty
        
        If Not IsEmpty(left) Then
            For i = LBound(left) To UBound(left)
                Arrays.Append arr, left(i)
            Next
        End If
        
        Arrays.Append arr, pivot
        
        If Not IsEmpty(right) Then
            For i = LBound(right) To UBound(right)
                Arrays.Append arr, right(i)
            Next
        End If
        
    End If
    
    QuickSort = arr
End Function
```

Example:
```basic
QuickSort arr, 1  ' By letter

'Input:  [[99,"D"],[1,"S"],[10,"P"],[79,"D"],[4,"H"],[38,"I"],[94,"I"],[40,"Z"],[16,"H"],[64,"E"],[41,"L"],[32,"T"],[20,"Q"],[58,"F"],[45,"C"],[26,"Y"],[37,"U"],[91,"I"],[62,"Q"],[9,"L"]]
'Output: [[45,"C"],[99,"D"],[79,"D"],[64,"E"],[58,"F"],[4,"H"],[16,"H"],[38,"I"],[94,"I"],[91,"I"],[41,"L"],[9,"L"],[10,"P"],[20,"Q"],[62,"Q"],[1,"S"],[32,"T"],[37,"U"],[26,"Y"],[40,"Z"]]


QuickSort arr, 0 ' By number
'Input: [[10,"C"],[12,"J"],[53,"A"],[54,"R"],[8,"W"],[67,"F"],[35,"M"],[70,"E"],[53,"Y"],[75,"C"],[46,"K"],[20,"N"],[9,"J"],[16,"P"],[9,"Y"],[27,"M"],[75,"X"],[67,"H"],[8,"H"],[32,"B"]]
'Output: [[8,"W"],[8,"H"],[9,"J"],[9,"Y"],[10,"C"],[12,"J"],[16,"P"],[20,"N"],[27,"M"],[32,"B"],[35,"M"],[46,"K"],[53,"A"],[53,"Y"],[54,"R"],[67,"F"],[67,"H"],[70,"E"],[75,"C"],[75,"X"]]
```


### Console output
Spacing out results with ``Tab()`` and ``Spc()``:
```VB
Debug.Print "ABC"; Tab(20); "DEF"; Tab(25); "GHI"
Debug.Print "ABC"; Spc(20); "DEF"; Spc(25); "GHI"
```
Outputs:
```
ABC                  DEF  GHI
ABC                    DEF                         GHI
```

### Dates
* Dates as variables can be entered in the following format: ``#yyyy/mm/dd#`` or ``#mm/dd/yyyy#``. Regardless of what method is used, VBA will automatically default to ``#mm/dd/yyyy#``.

### Dictionaries
* When pulling data from a table in a dictionary, erroneous cell values will not be entered; you need to perform an ``IsError(value)`` check and apply an alternative if true.

### Subroutines and Functions
#### Parameter arrays
```basic
Sub AnyNumberArgs(strName As String, ParamArray intScores() As Variant) 
 Dim intI As Integer 
 
 Debug.Print strName; " Scores" 
 ' Use UBound function to determine upper limit of array. 
 For intI = 0 To UBound(intScores()) 
 Debug.Print " "; intScores(intI) 
 Next intI 
End Sub

AnyNumberArgs "Jamie", 10, 26, 32, 15, 22, 24, 16 
AnyNumberArgs "Kelly", "High", "Low", "Average", "High"
```

### Miscellaneous
Remember when using the ``IIF`` function, that it will evaluate both parts of the _true_ and _false_ arguements. If there's a possiblity of returning an error from either of these, don't use this function.

#### The Static keyword
The ``Static`` keyword on a function variable 'remembers' what it contains even after exiting a function:

```basic
Function Records() as Dictionary
  Static data As Dictionary

  If data Is Nothing Then
    Set data = New Dictionary
    ' Read data
  Else
    ' Update data
  End If

  Set Records = data
End Function
```
On first calling `Records()`, the function first checks to see if there's anything assigned to `data` and if not, creates and assigns a dictionary object. On exiting the function, it returns the `data` object. Calling `Records()` a second time, the function remembers that `data` has already been assigned and updates and returns `data` instead.

You can use `Static` to create psuedo-objects. This function lets you keep track of the maximum values that gets passed to it. If `index` is -1, then it assumes you want to reset the lists. Otherwise, if `value` is `Empty`, then it returns whatever's at that index as an array of the code and index, otherwise it compares the value to what's in the array at index, and replaces if the new value is higher.
```basic
Private Function MaxValues(Optional index As Long = -1, Optional value As Variant = Empty, Optional code As String = "") As Variant
    Static values() As Variant
    Static codes() As Variant
    
    If index = -1 Then
        ReDim values(1 To QUARTERS)
        ReDim codes(1 To QUARTERS)
    Else
    
        If IsEmpty(value) Then
            MaxValues = Array(codes(index), values(index))
        Else
            If value > values(index) Then
                values(index) = value
                codes(index) = code
            End If
        End If
    End If
End Function

MaxValues                 ' Clears the data
MaxValues(1)(1)           ' Returns the code of whatever is at index 1
MaxValues 1, 5.7, "ABC"   ' Compares whats at index 1 with 5.7, and replaces if the value is higher
```

## Excel

### Data Tables
Reference|Range|Table
--|--|--
Relative|`A1`|`[FIELD]`
Absolute|`$A$1`|`[[FIELD]:[FIELD]]`



## Excel VBA
When plotting series, use a ``variant`` array to set the values. For missing data, use an ``Empty`` value and Excel will ignore plotting the point (Do not use ``Null`` as this can cause type mismatch errors).

### [Dictionary to JSON](https://github.com/myattadam/VBA-tools/blob/master/DictToJSON.bas)
This simple set of functions converts a nested dictionary structure to JSON format.
```VB
Sub saveJSON(filename As String, entity As Variant)
```
Saves the JSON file to the same location as the active workbook.

```VB
Function toJSON(entity As Variant) As String
```
A recursive function that converts a nested dictionary structure (and it's contents, whether they veriables, arrays, other dictionarys, etc.) to a JSON string.

A nested dictionary structure can be created as follows:
```VB
Dim lo As ListObject                  ' A ListObject is basically just an Excel table
Dim UID As String                     ' Our unique identifier

set record As Range                   ' Each record we're going to read is a row in our Excel table
Set data = New Dictionary             ' Where we're going to save our data

Set lo = Me.ListObjects(DATATABLE)    ' This example pulls records from a data table, but it could equally be via a query

For Each record In lo.DataBodyRange.Rows    ' Loops through 
    
    UID = Intersect(record, Me.Range(DATATABLE & "[UID]")).Value2
    
    If Not data.Exists(code) Then
        Set data(UID) = New Dictionary
            data(UID)("Code") = UID
            data(UID)("Client Name") = Intersect(record, Me.Range(DATATABLE & "[Client Name]")).Value2
        Set data(UID)("Transactions") = New Dictionary
            data(UID)("Total") = 0
    End If
                      
    period = Intersect(record, Me.Range(DATATABLE & "[Date]")).Value2
    
    data(code)("Transactions")(period) = Intersect(record, Me.Range(DATATABLE & "[Amount]")).Value2
    data(code)("Total") = data(code)("Total") + data(code)("Transactions")(period)
Next
```


### [datepicker.xlsm](https://github.com/myattadam/VBA-tools/blob/master/datepicker.xlsm)
I built this as a 'native' date picking tool as Excel 2010 doesn't come with a standard form control.
To use it in a workbook, copy the form and class to the workbooks project tree. To open the form, call _frmDatePicker.showForm Range("A1")_, replacing A1 with the destination cell the date needs to be written to.

_Note: Still needs a bit of work. I want to remove the hard coding for some of the values that set the colour of the buttons, etc._

### [shape_tools.ppam](https://github.com/myattadam/VBA-tools/blob/master/shape_tools.ppam)
This add-in currently contains a tool for copying and pasting just the size and position of a shape or range of shapes. The order in which you select the shapes to copy (shift click a shape to select more than one at a time) determines the order in which the size and positions are applied. Use this tool to quickly grab a layout of a slide and reapply it with a bunch of other shapes (including tables).
