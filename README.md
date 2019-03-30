# VBA tools and macro development

## General VBA

### Tips

* Dates as variables can be entered in the following format: ``#yyyy/mm/dd#`` or ``#mm/dd/yyyy#``. Regardless of what method is used, VBA will automatically default to ``#mm/dd/yyy#``.

* When pulling data from a table in a dictionary, erroneous cell values will not be entered; you need to perform an ``IsError(value)`` check and apply an alternative if true.

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
This simple set of functions converts a dictionary object to JSON format.
```VB
Sub saveJSON(filename As String, entity As Variant)
```
Saves the JSON file to the same location as the active workbook.

```VB
Function toJSON(entity As Variant) As String
```
A recursive function that converts a dictionary object (and it's contents, whether they veriables, arrays, other dictionarys, etc.) to a JSON string.

### [datepicker.xlsm](https://github.com/myattadam/VBA-tools/blob/master/datepicker.xlsm)
I built this as a 'native' date picking tool as Excel 2010 doesn't come with a standard form control.
To use it in a workbook, copy the form and class to the workbooks project tree. To open the form, call _frmDatePicker.showForm Range("A1")_, replacing A1 with the destination cell the date needs to be written to.

_Note: Still needs a bit of work. I want to remove the hard coding for some of the values that set the colour of the buttons, etc._
