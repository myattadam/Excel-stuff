# Excel tools and macro development

## VBA tricks

* When plotting a data series, use ``Empty`` to ignore plotting a point (Do not use ``Null`` as this can cause type mismatch errors)

* The ``Static`` keyword on a function variable 'remembers' what it contains even after exiting a function:

```VBA
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

## datepicker.xlsm
I built this as a 'native' date picking tool as Excel 2010 doesn't come with a standard form control.
To use it in a workbook, copy the form and class to the workbooks project tree. To open the form, call _frmDatePicker.showForm Range("A1")_, replacing A1 with the destination cell the date needs to be written to.

_Note: Still needs a bit of work. I want to remove the hard coding for some of the values that set the colour of the buttons, etc._
