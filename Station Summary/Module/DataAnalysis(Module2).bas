Attribute VB_Name = "Module2"
Sub MatchStation() ' Copy the stations from "Summary" to "Summary Sections"
Dim i As Integer
Dim j As Integer
i = 6
j = 3
Do While Not IsEmpty(Worksheets("Summary").Cells(1, i))
    Worksheets("Summary Sections").Cells(5, j).Value = Worksheets("Summary").Cells(1, i).Value
    i = i + 3
    j = j + 1
Loop
End Sub
Public Function ToColletter(Collet) 'Neccessary component for sorting. Converts columns from numbers to letter representations.
    ToColletter = Split(Cells(1, Collet).Address, "$")(1)
End Function
Sub SortYValues(i As Integer) 'sort each sations' values based on the y values
Dim col As String
Dim col1 As String
With Worksheets("Summary")
If i <> 0 Then                          'If i (input value) is not equal to 0, sort the station at column i
    col = ToColletter(i)                'convert column number to column letter (as "i : i+1" is not supported by "Range()")
    col1 = ToColletter(i + 1)
    .Columns(col & ":" & col1).Sort key1:=.Range(col & ":" & col), order1:=xlAscending, key2:=.Range(col1 & ":" & col1), order2:=xlDescending, Header:=xlYes
Else
    i = 5                                'If i = 0, sort entire worksheet
    Do While Not IsEmpty(Cells(1, i).Value)
        col = ToColletter(i)
        col1 = ToColletter(i + 1)
        .Columns(col & ":" & col1).Sort key1:=.Range(col & ":" & col), order1:=xlAscending, key2:=.Range(col1 & ":" & col1), order2:=xlDescending, Header:=xlYes
    i = i + 3
    Loop
End If
End With
End Sub
Sub SortY()         'To allow access to "SortYValue(0)" from the macro list since subs with input data are not listed
Call SortYValues(0)
End Sub
Sub SortZValues(i As Integer)               ' sort each station's values based on the z value in ascending order
Dim col As String
Dim col1 As String
With Worksheets("Summary")
If i <> 0 Then                              'If i (input value) is not equal to 0, sort the station at column i
    col = ToColletter(i)                    'convert column number to column letter (as "i : i+1" is not supported by "Range()")
    col1 = ToColletter(i + 1)
    .Columns(col & ":" & col1).Sort key1:=.Range(col1 & ":" & col1), order1:=xlAscending, key2:=.Range(col & ":" & col), order2:=xlDescending, Header:=xlYes
Else
    i = 5                                    'If i = 0, sort entire worksheet
    Do While Not IsEmpty(Cells(1, i).Value)
        col = ToColletter(i)
        col1 = ToColletter(i + 1)
        .Columns(col & ":" & col1).Sort key1:=.Range(col1 & ":" & col1), order1:=xlAscending, key2:=.Range(col & ":" & col), order2:=xlDescending, Header:=xlYes
    i = i + 3
    Loop
End If
End With
End Sub
Sub SortZ()         'To allow access to "SortZValue(0)" from the macro list since subs with input data are not listed
Call SortZValues(0)
End Sub
Sub TotalWidth()    'Calculate Top Width at each station
Dim i As Integer    'Column number (at each station)
Dim j As Integer    'Row number (at each station)
Dim sta As Integer  'the column number of the corresponding station on the "Summary Sections" Worksheet
Dim oneend As Double 'Smallest Y value at a station
Dim other As Double  'Largest Y value at a station
Dim twidth As Double 'Result total width
j = 2
i = 5               'starts from column 5, (the Y column of the first station)
sta = 3
Call SortYValues(0)                     'Sorts all stations on the worksheet based on Y value
Do While Not IsEmpty(Cells(1, i).Value) 'Finds the first and the last values in the Y column (the smallest and the biggest)
    With Worksheets("Summary")
    j = 2
    oneend = .Cells(j, i).Value
    Do While Not IsEmpty(.Cells(j, i).Value)
        j = j + 1
    Loop
    other = .Cells(j - 1, i).Value
    twidth = other - oneend
    End With
    Worksheets("Summary Sections").Cells(7, sta).Value = twidth 'Writes the result to the corresponding spot on "Summary Sections" Worksheet
i = i + 3
sta = sta + 1
Loop
End Sub
Sub TotalDepth()    'Calculate total depth at each station
Dim i As Integer    'Column number (at each station)
Dim j As Integer    'Row number (at each station)
Dim sta As Integer  'the column number of the corresponding station on the "Summary Sections" Worksheet
Dim oneend As Double 'The smallest Z value on a station
Dim other As Double  'The largest Z value on a station
Dim twidth As Double  'The result *Depth* of at a station (says width here because this code is essentially a copy of the TotalWidth sub)
j = 2
i = 6                 'Starts from column 6 (the Z column of the first station)
sta = 3
Call SortZValues(0)                     'Sorts all stations based on the Z values
Do While Not IsEmpty(Cells(1, i).Value) 'Finds the first and the last values in the Z column (the smallest and the biggest)
    With Worksheets("Summary")
    j = 2
    oneend = Cells(j, i).Value
    Do While Not IsEmpty(Cells(j, i).Value)
        j = j + 1
    Loop
    other = Cells(j - 1, i).Value
    twidth = other - oneend
    End With
    Worksheets("Summary Sections").Cells(6, sta).Value = twidth 'Wrties the result to the corresponding spot on the "Summary Sections" Worksheet
i = i + 3
sta = sta + 1
Loop
'Call SortYValues(0)        'initially included to sort the worksheet back (based on Y values) again. Not neccessary after the addition of the button
End Sub
Function Critical(ycol As Integer, Optional yesgunwale As Boolean = False) As Double 'This function loops through a station sorted based on Y values, finds the critcial value, then returns a corresponding row number) the functions needs 1 essential input values: ycol, the column number of a station. The second input value is optional. It determines whether the function needs to use user defined gunwale value
Dim i As Integer        'The row number variable
Dim y1 As Double        'Stores the Y value from previous (smaller) cell
Dim y2 As Double        'Stores the Y value of the current cell
Dim z1 As Double        'Stores the Z value from previous (smaller) cell
Dim z2 As Double        'Stores the Z value from current cell
Dim slope As Double     'Stores the resulting slope value
Dim piviot As Double    'The "Critical Value" of slope that separates Wall and Base
If Not IsEmpty(Worksheets("Summary Sections").Cells(3, 3).Value) Then
    Dim gunwale As Double   'The height of the gunwale. Needed so that the program can identify gunwale in its calculation
    piviot = Worksheets("Summary Sections").Cells(3, 3).Value   'Reads the value of pivot, which is stored in in cells(3,3) on "Summary Sections" worksheet
    gunwale = Worksheets("Summary Sections").Cells(15, 3).Value 'Reads the value of gunwale, which is stored in the "summary sections" worksheet
    i = 3   'starting at the third row, which holds the second value at a list (Note this function assumes the list has been sorted based on Y values. If it hasn't, call SortYValue before calling this function)
    With Worksheets("Summary")
    Do While Not IsEmpty(.Cells(i, ycol).Value) 'Looping through all the rows in a station
        y2 = .Cells(i, ycol).Value              'Y value of the current cell
        y1 = .Cells(i - 1, ycol).Value          'Y value of the previous cell
        z1 = .Cells(i, ycol + 1).Value           'Z value of the current cell
        z2 = .Cells(i - 1, ycol + 1).Value       'Z value of the previous cell
        If y2 - y1 <> 0 Then
            slope = (z2 - z1) / (y2 - y1)       'Calculate the slope
        Else
            slope = 99999                       'To avoid divide by 0 error
        End If
        'MsgBox (Abs(slope))
        If yesgunwale Then                      'If user specified gunwale value is needed (see the macro for the button on "Summary" worksheet for conditions)
            If Abs(slope) <= piviot And Abs(z2 - z1) < gunwale Then 'Compare the current slope with the pivot. To prevent the gunwale data from messing it up, the difference between the two cells has to be smaller than the height of gunwale
                Critical = i                 'returns the current row value modified a arbitray value to increase consistency
                Exit Function
            End If
        Else                                    'By default, as long as the slope is not 0 it should be fine
            If Abs(slope) <= piviot And z2 - z1 <> 0 Then
                Critical = i
                Exit Function
            End If
        End If
    i = i + 1
    Loop
    End With
    Critical = 0    'function returns zero if no result found
Else
    Critical = Worksheets("Summary Sections").Cells(2, 3).Value
End If
End Function
Sub BottomWallWidth()      'This sub calculates the bottom wall with based on the result from the Critcal function. It also doubles as the function that highlights and record pivot point values since it is the first function called by the button...
Dim ycol As Integer        'The column # of Y values at a station
Dim yrow As Integer        'The row # at which the pivot value is found. (probably makes more sense if I named it "CritRow". Oh well."
Dim crtval As Double       'The actual Y value of the pivot point
Dim i As Integer            'Apologize for my naming inconsistency... i here is the station column # on "Summary Sections" sheet
Dim width As Double         'The resulting width
i = 3
ycol = 5
Call SortYValues(0)
Do While Not IsEmpty(Worksheets("Summary").Cells(1, ycol))
    yrow = Critical(ycol)
    If yrow = 0 Then
        Worksheets("Summary Sections").Cells(8, i).Value = "None" 'writes the result on "Summary Sections"
    Else
        crtval = Cells(yrow, ycol).Value * 2
        '---------highlight/store crit values --------------
        Worksheets("Summary Sections").Cells(8, i).Value = Abs(crtval)
        Worksheets("Summary").Cells(yrow, ycol).Interior.Color = RGB(255, 0, 0)
        Worksheets("Summary Sections").Cells(17, i).Value = Worksheets("Summary").Cells(yrow, ycol).Value       'Writes Y critical value on row 17 "Summary Sections" sheet
        Worksheets("Summary Sections").Cells(18, i).Value = Worksheets("Summary").Cells(yrow, ycol + 1).Value   'Wrties Z critical value on row 18 "Summary Sections" sheet
        Worksheets("Summary Sections").Cells(19, i).Value = yrow                                                'writes the row number of the critical value.
        '---------------------------------------------------
    End If
i = i + 1
ycol = ycol + 3
Loop
End Sub
Sub WallandBaseHeight()     'This function calculates the wall and base height of the canoe based on the Critical function
Dim ycol As Integer         'The column # of Y values at a station
Dim yrow As Integer         'The row # at which the pivot value is found. (probably makes more sense if I named it "CritRow". Oh well."
Dim crtval As Double        'The actual Y value of the pivot point
Dim top As Double           'The Z value on the top of the canoe
Dim bottom As Double        'The Z value at the bottom of the canoe
Dim wallh As Double         'wall height
Dim baseh As Double         'base height
Dim i As Integer            'Apologize for my naming inconsistency... i here is the station column # on "Summary Sections" sheet
Dim j As Integer            'Used as a temporary variable to loop through each station column
Dim width As Double
i = 3
ycol = 5
Do While Not IsEmpty(Worksheets("Summary").Cells(1, ycol))

    Call SortZValues(ycol)                                     'Sorts the currrent station by its Z value
    j = 2
    bottom = Worksheets("Summary").Cells(j, ycol + 1).Value    'Since it is sorted in ascending order, the firs value is the smallest, which is the bottom
    Do While Not IsEmpty(Worksheets("Summary").Cells(j, ycol))
    j = j + 1
    Loop
    top = Worksheets("Summary").Cells(j - 1, ycol + 1).Value   'The last cell is the top
    
    
    Call SortYValues(ycol)                                      'Sorts the current station by its Y value
    yrow = Critical(ycol)                                       'Gets the critical point (row number)
    If yrow = 0 Then                                            '0 means the row is not found. If so, write "None" in the corresponding sections
        Worksheets("Summary Sections").Cells(9, i).Value = "None"
        Worksheets("Summary Sections").Cells(10, i).Value = "None"
    Else
        crtval = Worksheets("Summary").Cells(yrow, ycol + 1).Value  'gets the value of critical(pivot) point
        wallh = top - crtval                                        'Wall is the height from critcal point above
        baseh = crtval - bottom                                     'Base is the part from critcal point below
        'MsgBox ("crtval:" & crtval & " yrow:" & yrow & " top:" & top & " bottom:" & bottom) 'For testing purposes
        Worksheets("Summary Sections").Cells(9, i).Value = wallh     'writes the results
        Worksheets("Summary Sections").Cells(10, i).Value = baseh
    End If
i = i + 1
ycol = ycol + 3
Loop
End Sub
Sub BaseArea()          'this sub function calculates the base area based on the Critical Function
Dim dx As Double        'Difference in X value
Dim dy As Double
Dim dz As Double
Dim z1 As Double        'First z value (from previous cell)
Dim z2 As Double        'Second z value (curren cell)
Dim h As Double         'Height -- used in formula
Dim w As Double         'Width --- used in formula
Dim area As Double      'area -----result
Dim col As Integer      'colmn number (used in triangular approximation)
Dim icol As Integer     'initial column number  (used in triangular approximation)
col = 1
With Worksheets("Summary Sections")
Do While IsEmpty(.Cells(10, col)) 'initializing condition: first 'cell' reached
    col = col + 1
Loop
Do While .Cells(10, col).Value = "None" Or Cells(10, col).Value <> 0 'initializing condition: when it stops being 'None'
    col = col + 1
Loop
'---------------Triangular approxi for the first end----------------------
    dx = .Cells(5, col).Value - .Cells(5, 3).Value
    w = .Cells(8, col).Value
    area = dx * w / 2
    .Cells(11, col).Value = area
    col = col + 1
'---------------Approximation of area using trapezoids--------------------
Do While .Cells(10, col).Value <> "None" And .Cells(10, col).Value <> 0 And Not IsEmpty(.Cells(10, col))
    'z1 = .Cells(6, col - 1).Value - .Cells(10, col - 1).Value     #Ignorning the z displacement for now
    'z2 = .Cells(6, col).Value - .Cells(10, col).Value
    'dz = z2 - z1
    dx = .Cells(5, col).Value - .Cells(5, col - 1).Value
    w = .Cells(8, col - 1).Value + .Cells(8, col).Value
    area = dx * w / 2
    .Cells(11, col).Value = area
    col = col + 1
Loop
'------------------Trianglular approxi for the second end--------------------
icol = col - 1
Do While .Cells(10, col).Value = "None" 'initializing condition: when it stops being 'None'
    col = col + 1
Loop
col = col - 1
    dx = .Cells(5, col) - .Cells(5, icol)
    w = .Cells(8, icol)
    area = dx * w / 2
    .Cells(11, col).Value = area
End With
End Sub
