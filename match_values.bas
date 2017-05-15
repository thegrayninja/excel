Sub add_users_to_reporting()

x = 2
y = 6


rem cycling through sheet1 and matching it to values in sheet2
While Not Cells(x, 2).Value = "":
    strValue = Cells(x, y).Value
    lngRowNum = Application.Match(strValue, Sheet2.Range("A:A"), 0)
    If Not IsError(lngRowNum) Then
        Cells(x, 1).Value = lngRowNum
    Else: Cells(x, 1).Value = 1
    
    End If
    
    x = x + 1
Wend

rem since the above returns a number based on the row, we'll convert it to the actual value here
xx = 2

While Not Cells(xx, 2).Value = "":
    numvalue = Cells(xx, 1).Value
    Cells(xx, 1).Value = Sheet2.Cells(numvalue, 2).Value
    xx = xx + 1
    
    
Wend

End Sub
