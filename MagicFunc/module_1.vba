Public Sub datainput()
Dim n As Integer, m As Integer, i As Integer, xmas() As Single, ymas() As Single, a As Single
Cells(5, 2).Select
Do While ActiveCell.Value <> ""
n = n + 1: ActiveCell.Offset(0, 1).Select
Loop
Cells(6, 2).Select
Do While ActiveCell.Value <> ""
m = m + 1: ActiveCell.Offset(0, 1).Select
Loop
Range(Cells(8, 1), Cells(8, 10)).Select: Selection.Merge
If n = m Then
ActiveCell.Value = "Ooieoey caaaia a aeaa " & n & " yeaiaioia."
ReDim xmas(n) As Single
ReDim ymas(n) As Single
Cells(4, 1).Select
Cells(4, 1).Value = "#": ActiveCell.Offset(0, 1).Select
For i = 1 To m
ActiveCell.Value = i: ActiveCell.Offset(0, 1).Select
Next i
ActiveCell.Offset(0, -1).Select
Range(ActiveCell, ActiveCell.Offset(2, -n)).Select
Call borders
Cells(5, 2).Select
For i = 1 To n
xmas(i) = ActiveCell.Value: ActiveCell.Offset(0, 1).Select
Next i
Cells(6, 2).Select
For i = 1 To n
ymas(i) = ActiveCell.Value: ActiveCell.Offset(0, 1).Select
Next i
Cells(10, 1).Select

If Eeno1.Checkgraph.Value = True Then
Call graph(n)
End If

If Eeno1.checkinterpol.Value = True Then
If forminterpol.CheckBoxx.Value = False Then
Module2.checkxx
End If
 If forminterpol.Checklagr.Value = True Then
   Call Module2.interpol(xmas(), ymas(), n)
 End If
 If forminterpol.CheckNewton.Value = True Then
   Call Module2.newtoninterpol(xmas(), ymas(), n)
 End If
 End If
 
 If Eeno1.Checkapr.Value = True Then
 If formapr.Checkempir.Value = True Then
 Call aprempir(xmas(), ymas(), n)
 End If
End If

Else
ActiveCell.Value = "Eiee?anoai cia?aiee  O ia niioaaonoaoao eiee?anoao cia?aiee  Y!"
With Selection.Interior
.Color = vbRed
End With
End If
End Sub
Public Sub borders()
Selection.borders(xlDiagonalDown).LineStyle = xlNone
    Selection.borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.borders(xlEdgeLeft)
        .LineStyle = xlContinuous
    End With
    With Selection.borders(xlEdgeTop)
        .LineStyle = xlContinuous
    End With
    With Selection.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
    End With
    With Selection.borders(xlEdgeRight)
        .LineStyle = xlContinuous
    End With
    With Selection.borders(xlInsideVertical)
        .LineStyle = xlContinuous
    End With
    With Selection.borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
    End With
End Sub

Public Sub graph(n)
   ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Range(Cells(6, 2), Cells(6, n + 1))
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).XValues = Range(Cells(6, 2), Cells(6, n + 1))
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    With Selection
        .MarkerStyle = 1
        .MarkerSize = 5
    End With
    Selection.MarkerStyle = 8
    ActiveChart.SeriesCollection(1).Smooth = True
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
ActiveChart.Parent.Left = 8
ActiveChart.Parent.Top = 160
ActiveChart.Parent.Width = 400
ActiveChart.Parent.Height = 250
Cells(27, 1).Select
End Sub

