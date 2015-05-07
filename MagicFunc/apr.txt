Option Explicit
Public Sub aprempir(xmas() As Single, ymas() As Single, n)
Dim epsmas() As Variant, xar As Single, xgeom As Single, xgarm As Single, p1 As Integer, epsmin As Single, epskoord As Integer
Dim yar As Single, ygeom As Single, ygarm As Single, yar2 As Single, ygeom2 As Single, ygarm2 As Single, i As Integer
xar = (xmas(1) + xmas(n)) / 2
xgeom = Sqr(xmas(1) * xmas(n))
xgarm = (2 * xmas(1) * xmas(n)) / (xmas(1) + xmas(n))
yar = (ymas(1) + ymas(n)) / 2
ygeom = Sqr(ymas(1) * ymas(n))
ygarm = (2 * ymas(1) * ymas(n)) / (ymas(1) + ymas(n))

yar2 = 0
For i = 1 To n
If xmas(i) = xar Then
yar2 = ymas(i)
Exit For
End If
Next i
If yar2 = 0 Then
For i = 1 To n
If xar > xmas(i) Then
p1 = p1 + 1
End If
Next i
yar2 = ymas(p1) + (ymas(p1 + 1) - ymas(p1)) / (xmas(p1 + 1) - xmas(p1)) * (xar - xmas(p1))
End If

p1 = 0

ygeom2 = 0
For i = 1 To n
If xmas(i) = xgeom Then
ygeom2 = ymas(i)
Exit For
End If
Next i
If ygeom2 = 0 Then
For i = 1 To n
If xgeom > xmas(i) Then
p1 = p1 + 1
End If
Next i
ygeom2 = ymas(p1) + (ymas(p1 + 1) - ymas(p1)) / (xmas(p1 + 1) - xmas(p1)) * (xgeom - xmas(p1))
End If

p1 = 0

ygarm2 = 0
For i = 1 To n
If xmas(i) = xgarm Then
ygarm2 = ymas(i)
Exit For
End If
Next i
If ygarm2 = 0 Then
For i = 1 To n
If xgarm > xmas(i) Then
p1 = p1 + 1
End If
Next i
ygarm2 = ymas(p1) + (ymas(p1 + 1) - ymas(p1)) / (xmas(p1 + 1) - xmas(p1)) * (xgarm - xmas(p1))
End If
ReDim epsmas(2, 7) As Variant
epsmas(1, 1) = Abs(yar2 - yar): epsmas(2, 1) = "y = ax+b"
epsmas(1, 2) = Abs(yar2 - ygeom): epsmas(2, 2) = "y = a*b^x"
epsmas(1, 3) = Abs(yar2 - ygarm): epsmas(2, 3) = "y = 1/(ax+b)"
epsmas(1, 4) = Abs(ygeom2 - yar): epsmas(2, 4) = "y = a*ln(x)+b"
epsmas(1, 5) = Abs(ygeom2 - ygeom): epsmas(2, 5) = "y = a*x^b"
epsmas(1, 6) = Abs(ygarm2 - yar): epsmas(2, 6) = "y = a+b/x"
epsmas(1, 7) = Abs(ygarm2 - ygarm): epsmas(2, 7) = "y = x/(ax+b)"
epsmin = epsmas(1, 1): epskoord = 1
For i = 2 To 7
If epsmas(1, i) < epsmin Then
epsmin = epsmas(1, i): epskoord = i
End If
Next i
If formapr.Checkreshapr.Value = True Then
Range(ActiveCell, ActiveCell.Offset(0, 10)).Select: Selection.Merge: ActiveCell.Value = "Ii?aaaeaiea yiie?e?aneie oi?ioeu:": Selection.Font.Bold = True
ActiveCell.Offset(2, 0).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "X(a?.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = xar: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "X(aaii.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = xgeom: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "X(aa?i.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = xgarm: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "Y(a?.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = yar: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "Y(aaii.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = ygeom: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "Y(aa?i.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = ygarm: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "Y*(a?.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = yar2: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "Y*(aaii.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = ygeom2: ActiveCell.Offset(1, -2).Select
Range(ActiveCell, ActiveCell.Offset(0, 1)).Select: Selection.Merge: ActiveCell.Value = "Y*(aa?i.)=": ActiveCell.Offset(0, 1).Select: ActiveCell.Value = ygarm2
Range(ActiveCell, ActiveCell.Offset(-8, -1)).Select
Call Module1.borders
ActiveCell.Offset(10, 0).Select
For i = 1 To 7
ActiveCell.Value = "eps" & i: Selection.Font.Bold = True: ActiveCell.Offset(0, 1).Select
Next i
ActiveCell.Offset(1, -7).Select
For i = 1 To 7
ActiveCell.Value = epsmas(1, i): ActiveCell.Offset(0, 1).Select
Next i
ActiveCell.Offset(0, -1).Select: Range(ActiveCell, ActiveCell.Offset(-1, -6)).Select
Call Module1.borders
ActiveCell.Offset(3, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select: Selection.Merge
ActiveCell.Value = "eps(min)=eps" & epskoord & ", yiie?e?aneia au?a?aiea: " & epsmas(2, epskoord) & "."
ActiveCell.Offset(2, 0).Select
Else
Range(ActiveCell, ActiveCell.Offset(0, 8)).Select: Selection.Merge
ActiveCell.Value = "Yiie?e?aneay oi?ioea a iauai aeaa: " & epsmas(2, epskoord) & ".": Selection.Font.Bold = True
ActiveCell.Offset(2, 0).Select
End If
If formapr.Checkkoeflin.Value = True Then
If formapr.Checkreshapr.Value = False Then
Call aprkoeflin(xmas(), ymas(), n, epskoord)
Else
Call aprkoeflin(xmas(), ymas(), n, epskoord)
End If
End If
End Sub

Public Sub aprkoeflin(xmas() As Single, ymas() As Single, n, epskoord)
Dim zmas() As Single, qmas() As Single, i As Integer, por As Integer, delta1 As Single, delta2 As Single, delta As Single
Dim a0 As Single, a1 As Single, a As Single, b As Single
ReDim zmas(n) As Single
ReDim qmas(n) As Single
If epskoord = 1 Then
For i = 1 To n
zmas(i) = ymas(i): qmas(i) = xmas(i)
Next i
Call minor(zmas(), qmas(), n, a0, a1)
Range(ActiveCell, ActiveCell.Offset(0, 12)).Select: Selection.Merge: Selection.Font.Bold = True
ActiveCell.Value = "Au?a?aiea a eiie?aoiii aeaa: " & "y = " & a0 & "x +" & a1
ElseIf epskoord = 2 Then
For i = 1 To n
zmas(i) = Log(ymas(i)) / Log(10): qmas(i) = xmas(i)
Next i
Call minor(zmas(), qmas(), n, a0, a1)
a = 10 ^ a0: b = 10 ^ a1
Range(ActiveCell, ActiveCell.Offset(0, 12)).Select: Selection.Merge: Selection.Font.Bold = True
ActiveCell.Value = "Au?a?aiea a eiie?aoiii aeaa: " & "y = " & a & "*" & b & "^x"
ElseIf epskoord = 3 Then
For i = 1 To n
qmas(i) = xmas(i): zmas(i) = 1 / ymas(i)
Next i
Call minor(zmas(), qmas(), n, a0, a1)

Range(ActiveCell, ActiveCell.Offset(0, 12)).Select: Selection.Merge: Selection.Font.Bold = True
ActiveCell.Value = "Au?a?aiea a eiie?aoiii aeaa: " & "y = " & "1/(" & a0 & "x+" & a1 & ")"
ElseIf epskoord = 4 Then
For i = 1 To n
zmas(i) = ymas(i): qmas(i) = Log(xmas(i))
Next i
Call minor(zmas(), qmas(), n, a0, a1)

Range(ActiveCell, ActiveCell.Offset(0, 12)).Select: Selection.Merge: Selection.Font.Bold = True
ActiveCell.Value = "Au?a?aiea a eiie?aoiii aeaa: " & "y = " & a0 & "*ln(x)+" & a1
ElseIf epskoord = 5 Then
For i = 1 To n
zmas(i) = Log(ymas(i)) / Log(10): qmas(i) = Log(xmas(i)) / Log(10)
Next i
Call minor(zmas(), qmas(), n, a0, a1)
b = 10 ^ a0
Range(ActiveCell, ActiveCell.Offset(0, 12)).Select: Selection.Merge: Selection.Font.Bold = True
ActiveCell.Value = "Au?a?aiea a eiie?aoiii aeaa: " & "y = " & b & "*x^" & a1
ElseIf epskoord = 6 Then
For i = 1 To n
qmas(i) = 1 / xmas(i): zmas(i) = ymas(i)
Next i
Call minor(zmas(), qmas(), n, a0, a1)
Range(ActiveCell, ActiveCell.Offset(0, 12)).Select: Selection.Merge: Selection.Font.Bold = True
ActiveCell.Value = "Au?a?aiea a eiie?aoiii aeaa: " & "y = " & a & "+" & a1 & "/x"
Else
For i = 1 To n
zmas(i) = 1 / ymas(i): qmas(i) = 1 / xmas(i)
Next i
Call minor(zmas(), qmas(), n, a0, a1)
Range(ActiveCell, ActiveCell.Offset(0, 12)).Select: Selection.Merge: Selection.Font.Bold = True
ActiveCell.Value = "Au?a?aiea a eiie?aoiii aeaa: " & "y = x/(" & a0 & "x" & "+" & a1 & ")"
End If
ActiveCell.Offset(2, 0).Select
End Sub
Public Function suma(mas() As Single, n, p)
Dim i As Integer
For i = 1 To n
suma = suma + mas(i) ^ p
Next i
End Function
Public Function sumaxy(mas1() As Single, mas2() As Single, n, p)
Dim i As Integer
For i = 1 To n
sumaxy = sumaxy + mas1(i) * mas2(i) ^ p
Next i
End Function
Public Sub minor(zmas() As Single, qmas() As Single, n, a0, a1)
Dim delta1 As Single, delta2 As Single, delta As Single
delta1 = suma(zmas(), n, 1) * suma(qmas(), n, 2) - sumaxy(zmas(), qmas(), n, 1) * suma(qmas(), n, 1)
delta2 = n * sumaxy(zmas(), qmas(), n, 1) - suma(qmas(), n, 1) * suma(zmas(), n, 1)
delta = n * suma(qmas(), n, 2) - (suma(qmas(), n, 1)) ^ 2
a0 = delta1 / delta: a1 = delta2 / delta
End Sub
