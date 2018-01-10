Attribute VB_Name = "Module1"
Sub MakeStation() 'create stations from the Rhino output list
Dim i As Integer
Dim strow As Integer 'station row
Dim stcol As Integer  'station colomn
Dim col As String     'store colmn number in the form of letters (for the sorting data. Details see SortYValues sub in module 2)
Dim col1 As String    'store colmn number in the form of letters

strow = 2 'station row
stcol = 2 'station column
i = 2

'------Sort source data -----------------
With Worksheets("summary")
col = ToColletter(1)                'convert column number to column letter (as "i : i+1" is not supported by "Range()")
col1 = ToColletter(2)
col2 = ToColletter(3)
.Columns(col & ":" & col2).Sort key1:=.Range(col & ":" & col), order1:=xlAscending, key2:=.Range(col1 & ":" & col1), order2:=xlDescending, key3:=.Range(col2 & ":" & col2), order3:=xlDescending, Header:=xlYes
End With
'------Creat Stations --------------------
With Worksheets(1)
Do While Not IsEmpty(Cells(i, 1))
    x = Cells(i, 1).Value
    y = Cells(i, 2).Value
    Z = Cells(i, 3).Value
    If Cells(i - 1, 2).Value = y And Cells(i - 1, 3).Value = Z Then GoTo skip 'If the current cell is the same as the previous, skip
    If x <> Cells(i - 1, 1).Value Then
        stcol = stcol + 3
        strow = 2
        Cells(strow - 1, stcol).Value = "x =" 'Creating headers "x = #station"
        Cells(strow - 1, stcol + 1).Value = x
    End If
    Cells(strow, stcol).Value = y
    Cells(strow, stcol + 1).Value = Z
    strow = strow + 1
skip: i = i + 1
Loop
End With
End Sub
