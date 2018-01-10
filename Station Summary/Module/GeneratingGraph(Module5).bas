Attribute VB_Name = "Module5"
Sub CreateChart()

Dim cht As Object
  Set cht = ActiveSheet.Shapes.AddChart
  
  cht.Name = "Pivot"

'Determine the chart type
  cht.Chart.ChartType = xlXYScatter

'Delete all uneccessary series
  For Each s In cht.Chart.SeriesCollection
  s.Delete
  Next s

End Sub

Sub LocatePivot(Optional displacement As Integer = -1, Optional position As Integer = 0)
Dim i As Integer
Dim j As Integer
Dim k As Integer
k = 2
j = 2
i = 6
With Worksheets("Summary")
Do While Not IsEmpty(.Cells(1, i)) And Not .Cells(1, i) = 0   'find station 0 on "Summary"
i = i + 3
Loop

i = i + position * 3

If position <> 0 Then 'If there is a valid position value
    i = position
End If
Do While Not IsEmpty(.Cells(j, i - 1))
    If .Cells(j, i - 1).Interior.Color = RGB(255, 0, 0) Then
        Exit Do
    End If
j = j + 1
Loop
Do While Not IsEmpty(.Cells(k, i))
k = k + 1
Loop
End With

For Each s In ActiveWorkbook.Worksheets("Summary Sections").ChartObjects("Pivot").Chart.SeriesCollection 'Delete existing series
  s.Delete
  Next s
Worksheets("Summary").Activate
With ActiveWorkbook.Worksheets("Summary Sections").ChartObjects("Pivot").Chart.SeriesCollection.NewSeries
    .Values = Range(Cells(2, i), Cells(k - 1, i))
    .XValues = Range(Cells(2, i - 1), Cells(k - 1, i - 1))
    .Name = "YvsZ"
End With
Worksheets("Summary Sections").Activate

ActiveWorkbook.Sheets("Summary Sections").ChartObjects("Pivot").Chart.SeriesCollection("YvsZ").Points(j - 1).MarkerBackgroundColor = RGB(255, 0, 0)
ActiveWorkbook.Sheets("Summary Sections").ChartObjects("Pivot").Chart.SeriesCollection("YvsZ").Points(j - 1).MarkerForegroundColor = RGB(0, 0, 0)
End Sub

Sub LocateLine() 'Side view of LocatePivot
Dim i As Integer
Dim j As Integer
Dim first As Integer

With Worksheets("Summary Sections")
i = 3
Do While IsEmpty(.Cells(17, i))
i = i + 1
Loop
first = i
Do While Not IsEmpty(.Cells(17, i))
i = i + 1
Loop
i = i - 1
End With

For Each s In ActiveWorkbook.Sheets("Summary Sections").ChartObjects("Line").Chart.SeriesCollection
  s.Delete
  Next s
With ActiveWorkbook.Sheets("Summary Sections").ChartObjects("Line").Chart.SeriesCollection.NewSeries
    .Values = Worksheets("Summary").Range("C:C")
    .XValues = Worksheets("Summary").Range("A:A")
    .Name = "XvsZ"
With ActiveWorkbook.Sheets("Summary Sections").ChartObjects("Line").Chart.SeriesCollection.NewSeries
    .Values = Worksheets("Summary Sections").Range(Cells(18, first).Address & ":" & Cells(18, i).Address)
    .XValues = Worksheets("Summary Sections").Range(Cells(5, first).Address & ":" & Cells(5, i).Address)
    .Name = "Line"
End With

End With

End Sub

Sub MovePivotGraph()
Dim position As Integer

For Each cell In Selection
position = cell.Column
Exit For
Next cell

position = (position - 3) * 3 + 6

Call LocatePivot(-1, position)
End Sub
