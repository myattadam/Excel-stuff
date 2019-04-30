# VBA tools and macro development

## General VBA

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

### Miscellaneous
* Remember when using the ``IIF`` function, that it will evaluate both parts of the _true_ and _false_ arguements. If there's a possiblity of returning an error from either of these, don't use this function.

* The ``Static`` keyword on a function variable 'remembers' what it contains even after exiting a function:

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

On first calling ``Records()``, the function first checks to see if there's anything assigned to ``data`` and if not, creates and assigns a dictionary object. On exiting the function, it returns the ``data`` object. Calling ``Records()`` a second time, the function remembers that ``data`` has already been assigned and updates and returns ``data`` instead.

## Excel VBA

### Tips
* \[Excel\] When plotting series, use a ``variant`` array to set the values. For missing data, use an ``Empty`` value and Excel will ignore plotting the point (Do not use ``Null`` as this can cause type mismatch errors).

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
