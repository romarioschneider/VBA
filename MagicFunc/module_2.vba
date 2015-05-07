Option Explicit
Dim deltamas() As Single, x As Single
Public Sub checkxx()
If forminterpol.Checklagr.Value = False And forminterpol.CheckNewton = False Then
Else
If forminterpol.CheckBoxx.Value = False Then
x = Application.InputBox("Aaaaeoa cia?aiea O aey ii?aaaeaiey cia?aiey ooieoee:", "/X/...", , , , , , 1)
End If
End If
End Sub
Public Sub interpol(xmas() As Single, ymas() As Single, n)
Dim i As Integer, ii As Integer, iii As Integer, chis() As Single, znam() As Single, y As Single
Dim s1 As Integer, s2 As Integer, q As Integer
If forminterpol.CheckBoxx.Value = True Then
x = Application.InputBox("Aaaaeoa cia?aiea O aey ii?aaaeaiey cia?aiey ooieoee:", "/X/...", , , , , , 1)
End If
If forminterpol.Checkreshinterpol.Value = False Then
If forminterpol.bandinterpol.Value = False Then

ReDim chis(n) As Single
ReDim znam(n) As Single

For i = 1 To n
chis(i) = 1
For ii = 1 To n
If ii <> i Then
chis(i) = chis(i) * (x - xmas(ii))
End If
Next ii
Next i

For i = 1 To n
znam(i) = 1
For ii = 1 To n
If ii <> i Then
znam(i) = znam(i) * (xmas(i) - xmas(ii))
End If
Next ii
Next i

For i = 1 To n
y = y + (chis(i) / znam(i) * ymas(i))
Next i
Range(ActiveCell, ActiveCell.Offset(0, 8)).Select: Selection.Merge
ActiveCell.Value = "?aoaiea ii Eaa?ai?o: i?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select

Else

Do While forminterpol.point1interpol.text = ""
q = MsgBox("Ia caaai ia?aeuiue ocae i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/Ia?aeuiue ocae/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While forminterpol.point2interpol.text = ""
q = MsgBox("Ia caaai eiia?iue ocae i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/Eiia?iue ocae/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol) > Val(forminterpol.point2interpol))
q = MsgBox("Iiia? ia?aeuiiai ocea aieuoa iiia?a eiia?iiai ocea i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) = Int(Val(forminterpol.point2interpol))
q = MsgBox("Iiia? ia?aeuiiai ocea ?aaai iiia?o eiia?iiai ocea i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) < 1
q = MsgBox("Iiia? ia?aeuiiai ocea ia ii?ao auou iaiuoa 1-u! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point2interpol)) > n
q = MsgBox("Iiia? eiia?iiai ocea auoiaeo ca i?aaaeu iauaai eiee?anaa oceia! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point2interpol)) <= 0
q = MsgBox("Iiia? eiia?iiai ocea ia ii?ao auou eee ?aaiyouny 0-?! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) <= 0
q = MsgBox("Iiia? ia?aeuiiai ocea ia ii?ao auou eee ?aaiyouny 0-?! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop





s1 = Int(Val(forminterpol.point1interpol))
s2 = Int(Val(forminterpol.point2interpol))

ReDim chis(s2 - s1 + 1) As Single
ReDim znam(s2 - s1 + 1) As Single
q = 0

For i = s1 To s2
q = q + 1
chis(q) = 1
For ii = s1 To s2
If ii <> i Then
chis(q) = chis(q) * (x - xmas(ii))
End If
Next ii
Next i

q = 0

For i = s1 To s2
q = q + 1
znam(q) = 1
For ii = s1 To s2
If ii <> i Then
znam(q) = znam(q) * (xmas(i) - xmas(ii))
End If
Next ii
Next i

q = 0

For i = s1 To s2
q = q + 1
y = y + (chis(q) / znam(q) * ymas(i))
Next i
Range(ActiveCell, ActiveCell.Offset(0, 8)).Select: Selection.Merge
ActiveCell.Value = "?aoaiea ii Eaa?ai?o aey aeaiaciia #" & s1 & " - #" & s2 & ":i?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select

End If

Else

If forminterpol.bandinterpol.Value = False Then
Dim chistext() As String, znamtext() As String, ytext() As Single

ReDim chis(n) As Single
ReDim znam(n) As Single
ReDim chistext(n) As String
ReDim znamtext(n) As String
ReDim ytext(n) As Single

For i = 1 To n
ytext(i) = ymas(i)
Next i

For i = 1 To n
chis(i) = 1
For ii = 1 To n
If ii <> i Then
chis(i) = chis(i) * (x - xmas(ii))
chistext(i) = chistext(i) & "(" & x & "-" & xmas(ii) & ")"
End If
Next ii
Next i

For i = 1 To n
znam(i) = 1
For ii = 1 To n
If ii <> i Then
znam(i) = znam(i) * (xmas(i) - xmas(ii))
znamtext(i) = znamtext(i) & "(" & xmas(i) & "-" & xmas(ii) & ")"
End If
Next ii
Next i
ActiveCell.Value = "P" & n - 1 & "(" & x & ") = ": ActiveCell.Offset(1, 0).Select

For i = 1 To n
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
If i = 1 Then
ActiveCell.Value = " " & " = " & chistext(i) & " / " & znamtext(i) & " * " & ytext(i) & " +"
ActiveCell.Offset(1, 0).Select
Else
If i <> n Then
ActiveCell.Value = chistext(i) & " / " & znamtext(i) & " * " & ymas(i) & " +"
ActiveCell.Offset(1, 0).Select
Else
ActiveCell.Value = chistext(i) & " / " & znamtext(i) & " * " & ymas(i) & " ="
ActiveCell.Offset(1, 0).Select
End If
End If
Next i




For i = 1 To n
ytext(i) = (chis(i) / znam(i) * ymas(i))
y = y + ytext(i)
Next i

Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
For i = 1 To n
If i = 1 Then
ActiveCell.Value = " " & " = " & ytext(i)
Else
If ytext(i) < 0 Then
ActiveCell.Value = ActiveCell.Value & " - " & Abs(ytext(i))
Else
ActiveCell.Value = ActiveCell.Value & " + " & ytext(i)
End If
End If
Next i
ActiveCell.Value = ActiveCell.Value & " " & " = " & y: ActiveCell.Offset(2, 0).Select
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select
ActiveCell.Value = "?aoaiea ii Eaa?ai?o: i?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select

Else

Dim chistext2() As String, znamtext2() As String, ytext2() As Single

Do While forminterpol.point1interpol.text = ""
q = MsgBox("Ia caaai ia?aeuiue ocae i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/Ia?aeuiue ocae/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While forminterpol.point2interpol.text = ""
q = MsgBox("Ia caaai eiia?iue ocae i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/Eiia?iue ocae/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Val(forminterpol.point1interpol) > Val(forminterpol.point2interpol)
q = MsgBox("Iiia? ia?aeuiiai ocea aieuoa iiia?a eiia?iiai ocea i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Val(forminterpol.point1interpol) = Val(forminterpol.point2interpol)
q = MsgBox("Iiia? ia?aeuiiai ocea ?aaai iiia?o eiia?iiai ocea i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Val(forminterpol.point1interpol) < 1
q = MsgBox("Iiia? ia?aeuiiai ocea ia ii?ao auou iaiuoa 1-u! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Val(forminterpol.point2interpol) > n
q = MsgBox("Iiia? eiia?iiai ocea auoiaeo ca i?aaaeu iauaai eiee?anaa oceia! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point2interpol)) <= 0
q = MsgBox("Iiia? eiia?iiai ocea ia ii?ao auou eee ?aaiyouny 0-?! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) <= 0
q = MsgBox("Iiia? ia?aeuiiai ocea ia ii?ao auou eee ?aaiyouny 0-?! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop


s1 = Int(Val(forminterpol.point1interpol))
s2 = Int(Val(forminterpol.point2interpol))
ReDim chistext2(s2 - s1 + 1) As String
ReDim znamtext2(s2 - s1 + 1) As String
ReDim ytext2(s2 - s1 + 1) As Single


ReDim chis(s2 - s1 + 1) As Single
ReDim znam(s2 - s1 + 1) As Single
q = 0
For i = s1 To s2
q = q + 1
ytext2(q) = ymas(q)
Next i
q = 0

For i = s1 To s2
q = q + 1
chis(q) = 1
For ii = s1 To s2
If ii <> i Then
chis(q) = chis(q) * (x - xmas(ii))
chistext2(q) = chistext2(q) & "(" & x & "-" & xmas(ii) & ")"
End If
Next ii
Next i

q = 0

For i = s1 To s2
q = q + 1
znam(q) = 1
For ii = s1 To s2
If ii <> i Then
znam(q) = znam(q) * (xmas(i) - xmas(ii))
znamtext2(q) = znamtext2(q) & "(" & xmas(i) & "-" & xmas(ii) & ")"
End If
Next ii
Next i

q = 0
ActiveCell.Value = "P" & s2 - s1 & "(" & x & ") = ": ActiveCell.Offset(1, 0).Select

For i = s1 To s2
q = q + 1
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
If i = s1 Then
ActiveCell.Value = " " & " = " & chistext2(q) & " / " & znamtext2(q) & " * " & ytext2(q) & " +"
ActiveCell.Offset(1, 0).Select
Else
If i <> s2 Then
ActiveCell.Value = chistext2(q) & " / " & znamtext2(q) & " * " & ymas(q) & " +"
ActiveCell.Offset(1, 0).Select
Else
ActiveCell.Value = chistext2(q) & " / " & znamtext2(q) & " * " & ymas(q) & " ="
ActiveCell.Offset(1, 0).Select
End If
End If
Next i
q = 0
For i = s1 To s2
q = q + 1
ytext2(q) = (chis(q) / znam(q) * ymas(i))
y = y + ytext2(q)
Next i
q = 0
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
For i = s1 To s2
q = q + 1
If i = s1 Then
ActiveCell.Value = " " & " = " & ytext2(q)
Else
If ytext2(q) < 0 Then
ActiveCell.Value = ActiveCell.Value & " - " & Abs(ytext2(q))
Else
ActiveCell.Value = ActiveCell.Value & " + " & ytext2(q)
End If
End If
Next i
ActiveCell.Value = ActiveCell.Value & " " & " = " & y: ActiveCell.Offset(2, 0).Select
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
ActiveCell.Value = "?aoaiea ii Eaa?ai?o aey aeaiaciia #" & s1 & " - #" & s2 & ":i?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select
End If

End If

End Sub

Public Sub deltatable(xmas() As Single, ymas() As Single, d)
Dim i As Integer, ii As Integer, z As Integer
ReDim deltamas(d, d) As Single

For i = 1 To d
deltamas(i, 1) = ymas(i)
Next i
z = d - 1
For i = 2 To d + 1
For ii = 1 To z
deltamas(ii, i) = deltamas(ii + 1, i - 1) - deltamas(ii, i - 1)
Next ii
z = z - 1
Next i
Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oaaeeoa eiia?iuo ?aciinoae:": ActiveCell.Offset(1, 0).Select
ActiveCell.Value = "X": ActiveCell.Offset(0, 1).Select
ActiveCell.Value = "Y": ActiveCell.Offset(0, 1).Select
For i = 1 To d - 1
ActiveCell.Value = "delta" & i: ActiveCell.Offset(0, 1).Select
Next i
ActiveCell.Offset(1, -d - 1).Select
For i = 1 To d
ActiveCell.Value = xmas(i): ActiveCell.Offset(1, 0).Select
Next i
ActiveCell.Offset(-d, 1).Select
For i = 1 To d
For ii = 1 To d
ActiveCell.Value = deltamas(ii, i): ActiveCell.Offset(1, 0).Select
Next ii
ActiveCell.Offset(-d, 1).Select
Next i
ActiveCell.Offset(-1, -1).Select: Range(ActiveCell, ActiveCell.Offset(0, -d)).Select: Selection.Font.Bold = True: ActiveCell.Select
Range(ActiveCell, ActiveCell.Offset(d, d)).Select
Call Module1.borders
ActiveCell.Offset(d + 3, 0).Select
End Sub

Public Sub newtoninterpol(xmas() As Single, ymas() As Single, n)
Dim i As Integer, ii As Integer, h As Single, p1 As Integer, p2 As Integer, teleport As Integer, q As Integer
Dim y As Single, controlh() As Single, dev As Single, devetalon As Single, xmasb() As Single, ymasb() As Single

If forminterpol.bandinterpol.Value = True Then
Dim s1 As Single, s2 As Single

Do While forminterpol.point1interpol.text = ""
q = MsgBox("Ia caaai ia?aeuiue ocae i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/Ia?aeuiue ocae/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While forminterpol.point2interpol.text = ""
q = MsgBox("Ia caaai eiia?iue ocae i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/Eiia?iue ocae/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) > Int(Val(forminterpol.point2interpol))
q = MsgBox("Iiia? ia?aeuiiai ocea aieuoa iiia?a eiia?iiai ocea i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) = Int(Val(forminterpol.point2interpol))
q = MsgBox("Iiia? ia?aeuiiai ocea ?aaai iiia?o eiia?iiai ocea i?iia?ooea! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) < 1
q = MsgBox("Iiia? ia?aeuiiai ocea ia ii?ao auou iaiuoa 1-u! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point2interpol)) > n
q = MsgBox("Iiia? eiia?iiai ocea auoiaeo ca i?aaaeu iauaai eiee?anaa oceia! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point2interpol)) <= 0
q = MsgBox("Iiia? eiia?iiai ocea ia ii?ao auou eee ?aaiyouny 0-?! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

Do While Int(Val(forminterpol.point1interpol)) <= 0
q = MsgBox("Iiia? ia?aeuiiai ocea ia ii?ao auou eee ?aaiyouny 0-?! Eciaieou aaiiua?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If q = vbYes Then
forminterpol.Show
Else
Exit Sub
End If
Loop

s1 = Int(Val(forminterpol.point1interpol))
s2 = Int(Val(forminterpol.point2interpol))
ReDim xmasb(s2 - s1 + 1) As Single
ReDim ymasb(s2 - s1 + 1) As Single
For i = s1 To s2
q = q + 1
xmasb(q) = xmas(i): ymasb(q) = ymas(i)
Next i
q = s2 - s1 + 1

Else

ReDim xmasb(n) As Single
ReDim ymasb(n) As Single
For i = 1 To n
xmasb(i) = xmas(i): ymasb(i) = ymas(i)
Next i

q = n
End If





ReDim controlh(q - 1) As Single

For i = 1 To q - 1
controlh(i) = xmasb(i + 1) - xmasb(i)
Next i

For i = 1 To q - 2
If controlh(i) <> controlh(i + 1) Then
teleport = MsgBox("Iaia?o?aia iaoi?iay ?aaiiioaae?iiinou oceia! I?iaie?eou?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua/...")
If teleport = vbYes Then
teleport = 10
Exit For
Else
Exit Sub
End If
End If
Next i

If teleport = 10 Then
devetalon = Application.InputBox("Aaaaeoa iaiaoiaeio? OI?IO? ?aaiiioaae?iiinou:", "/Oi?iay ?aaiiioaae?iiinou/...", controlh(1), , , , , 1)
dev = Application.InputBox("Aaaaeoa iaeneiaeuii aicii?iia ioeeiiaiea aey ?aaiiioaae?iiinoe." & vbNewLine & "Aey cia?aiey, ia ioee?a?uaainy io 0 cai?aiea oaaa aoaao ?aaii" & vbNewLine & "n?aaiaio a?eoiaoe?aneiio cia?aie? ?acieo ia?ao oceaie.", "/Ioeeiiaiea~/...", 0, , , , , 1)

For i = 1 To q - 1
If Abs(Abs(devetalon) - controlh(i)) > Abs(dev) Then
teleport = MsgBox("Eioa?aae ia?ao oceii #" & i & " e oceii #" & i + 1 & " i?aauoaao iaeneiaeuiia ioeeiiaiea. I?iaie?eou?", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/...")
If teleport = vbYes Then
Else
Exit Sub
End If
End If
Next i

h = 0

If dev <> 0 Then
For i = 1 To q - 1
h = h + controlh(i)
Next i
h = h / (q - 1)
Else
MsgBox "Iaeneiaeuiia ioeeiiaiea ?aaii 0. I?iaie?aiea iaaicii?ii!", vbCritical + vbYesNo, "/I?iaa?uoa aaiiua!/"
Exit Sub
End If

Else

h = xmasb(2) - xmasb(1)
End If

If forminterpol.CheckBoxx.Value = True Then
x = Application.InputBox("Aaaaeoa cia?aiea O aey ii?aaaeaiey cia?aiey ooieoee:", "/X/...", , , , , , 1)
End If



For i = 1 To q
If x >= xmasb(i) Then
p1 = p1 + 1
End If
Next i

For i = q To 1 Step -1
If x <= xmasb(i) Then
p2 = p2 + 1
End If
Next i

If p1 = 0 Then
Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
If forminterpol.bandinterpol.Value = False Then
ActiveCell.Value = "?aoaiea ii Iu?oiio aey anae oaaeeou:": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ca i?aaaeaie oaaeeou (ia?aa oceii #1)."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #1:": ActiveCell.Offset(2, 0).Select
Else
ActiveCell.Value = "?aoaiea ii Iu?oiio aey aeaiaciia #" & s1 & "-#" & s2 & ":": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ca i?aaaeaie aeaiaciia (ia?aa oceii #" & s1 & ")."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #1:": ActiveCell.Offset(2, 0).Select
End If

If forminterpol.Checkreshinterpol.Value = False Then
Call deltatable(xmasb(), ymasb(), q)
Call formula1(xmasb(), ymasb(), q, y, x, h)
Else
Call deltatable(xmasb(), ymasb(), q)
Call formula1show(xmasb(), ymasb(), q, y, x, h)
End If


Else


If p1 < p2 Then
If forminterpol.bandinterpol.Value = False Then
Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "?aoaiea ii Iu?oiio aey anae oaaeeou:": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ia?ao oceii #" & p1 & " e oceii #" & p1 + 1 & "."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #1:": ActiveCell.Offset(2, 0).Select
Else
ActiveCell.Value = "?aoaiea ii Iu?oiio aey aeaiaciia #" & s1 & "-#" & s2 & ":": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ia?ao oceaie #" & p1 & " e oceii #" & p1 + 1 & " aeaiaciia."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #1:": ActiveCell.Offset(2, 0).Select
End If
If forminterpol.Checkreshinterpol.Value = False Then
Call deltatable(xmasb(), ymasb(), q)

Call formula1(xmasb(), ymasb(), q, y, x, h)


Else
Call deltatable(xmasb(), ymasb(), q)
Call formula1show(xmasb(), ymasb(), q, y, x, h)
End If
Else

If p1 > p2 And p2 <> 0 Then
If forminterpol.bandinterpol.Value = False Then
Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "?aoaiea ii Iu?oiio aey anae oaaeeou:": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ia?ao oceii #" & p1 & " e oceii #" & p1 + 1 & "."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #2:": ActiveCell.Offset(2, 0).Select
Else
ActiveCell.Value = "?aoaiea ii Iu?oiio aey aeaiaciia #" & s1 & "-#" & s2 & ":": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ia?ao oceaie #" & p1 & " e  #" & p1 + 1 & " aeaiaciia."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #2:": ActiveCell.Offset(2, 0).Select
End If

If forminterpol.Checkreshinterpol.Value = False Then
Call deltatable(xmasb(), ymasb(), q)
Call formula2(xmasb(), ymasb(), q, y, x, h)
Else
Call deltatable(xmasb(), ymasb(), q)
Call formula2show(xmasb(), ymasb(), q, y, x, h)
End If

Else

Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
If forminterpol.bandinterpol.Value = False Then
ActiveCell.Value = "?aoaiea ii Iu?oiio aey anae oaaeeou:": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ca i?aaaeaie oaaeeou (iinea ocea #" & q & ")."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #2:": ActiveCell.Offset(2, 0).Select
Else
ActiveCell.Value = "?aoaiea ii Iu?oiio aey aeaiaciia #" & s1 & "-#" & s2 & ":": Selection.Font.Bold = True: ActiveCell.Offset(2, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "Oeacaiiue aey ?an??oa ocae X=" & x & " iaoiaeony ca i?aaaeaie oaaeeou (iinea ocea #" & s2 & ")."
ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Eniieucoai eioa?iieyoeiiio? oi?ioeo #2:": ActiveCell.Offset(2, 0).Select
End If

If forminterpol.Checkreshinterpol.Value = False Then
Call deltatable(xmasb(), ymasb(), q)
Call formula2(xmasb(), ymasb(), q, y, x, h)
Else
Call deltatable(xmasb(), ymasb(), q)
Call formula2show(xmasb(), ymasb(), q, y, x, h)
End If

End If
End If
End If

End Sub

Public Sub formula1(xmas() As Single, ymas() As Single, n, y, x, h)
Dim i As Integer, ii As Integer, mnx As Single, factorial As Single
y = ymas(1): factorial = 1: mnx = 1
For i = 1 To n - 1
factorial = factorial * i
mnx = mnx * (x - xmas(i))
y = y + deltamas(1, i + 1) / (h ^ i * factorial) * mnx
Next i
ActiveCell.Offset(1, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "I?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select
End Sub

Public Sub formula2(xmas() As Single, ymas() As Single, n, y, x, h)
Dim i As Integer, ii As Integer, factorial As Single, mnx As Single
y = ymas(n): factorial = 1: mnx = 1
For i = n To 2 Step -1
factorial = factorial * (n - i + 1)
mnx = mnx * (x - xmas(i))
y = y + deltamas(i - 1, n - i + 2) / (h ^ (n - i + 1) * factorial) * mnx
Next i
ActiveCell.Offset(1, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "I?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select
End Sub

Public Sub formula1show(xmas() As Single, ymas() As Single, n, y, x, h)
Dim i As Integer, ii As Integer, mnx As Single, factorial As Single
Dim mnxtext As String, textmas() As Single
ReDim textmas(n - 1) As Single
y = ymas(1): factorial = 1: mnx = 1
ActiveCell.Value = "P" & n & "(" & x & ") = " & y & " + ": ActiveCell.Offset(1, 0).Select
For i = 1 To n - 1
factorial = factorial * i
mnx = mnx * (x - xmas(i)): mnxtext = mnxtext & "(" & x & "-" & xmas(i) & ")"
textmas(i) = deltamas(1, i + 1) / (h ^ i * factorial) * mnx
y = y + textmas(i)
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
If i <> n - 1 Then
If i = 1 Then
ActiveCell.Value = "+ " & deltamas(1, i + 1) & " / " & h & " * " & mnxtext & " +"
Else
ActiveCell.Value = "+ " & deltamas(1, i + 1) & " / (" & h & "^" & i & "*" & factorial & ")" & " * " & mnxtext & " +"
End If
Else
ActiveCell.Value = "+ " & deltamas(1, i + 1) & " / (" & h & "^" & i & "*" & factorial & ")" & " * " & mnxtext & " ="
End If
ActiveCell.Offset(1, 0).Select
Next i
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
ActiveCell.Value = "" & " = " & ymas(1)
For i = 1 To n - 1
If textmas(i) > 0 Then
ActiveCell.Value = ActiveCell.Value & " + " & textmas(i)
Else
If textmas(i) = 0 Then
Else
ActiveCell.Value = ActiveCell.Value & " - " & Abs(textmas(i))
End If
End If
Next i
ActiveCell.Value = ActiveCell.Value & "" & " = " & y
ActiveCell.Offset(1, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "I?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select
End Sub
Public Sub formula2show(xmas() As Single, ymas() As Single, n, y, x, h)
Dim i As Integer, q As Integer, factorial As Single, mnx As Single
Dim mnxtext As String, textmas() As Single
ReDim textmas(n) As Single
y = ymas(n): factorial = 1: mnx = 1
ActiveCell.Value = "P" & n & "(" & x & ") = " & ymas(n) & " +": ActiveCell.Offset(1, 0).Select
For i = n To 2 Step -1
q = q + 1
factorial = factorial * (n - i + 1)
mnx = mnx * (x - xmas(i)): mnxtext = mnxtext & "(" & x & "-" & xmas(i) & ")"
textmas(q) = deltamas(i - 1, n - i + 2) / (h ^ (n - i + 1) * factorial) * mnx
y = y + textmas(q)
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
If i <> 2 Then
If i = n Then
ActiveCell.Value = "+ " & deltamas(i - 1, n - i + 2) & " / " & h & " * " & mnxtext & " +"
Else
ActiveCell.Value = "+ " & deltamas(i - 1, n - i + 2) & " / (" & h & "^" & n - 1 + 1 & "*" & factorial & ")" & " * " & mnxtext & " +"
End If
Else
ActiveCell.Value = "+ " & deltamas(i - 1, n - i + 2) & " / (" & h & "^" & n - 1 + 1 & "*" & factorial & ")" & " * " & mnxtext & " ="
End If
ActiveCell.Offset(1, 0).Select
Next i
Range(ActiveCell, ActiveCell.Offset(0, 254)).Select: Selection.Merge
ActiveCell.Value = "" & " = " & ymas(n)
For i = 1 To n - 1
If textmas(i) > 0 Then
ActiveCell.Value = ActiveCell.Value & " + " & textmas(i)
Else
If textmas(i) = 0 Then
Else
ActiveCell.Value = ActiveCell.Value & " - " & Abs(textmas(i))
End If
End If
Next i
ActiveCell.Value = ActiveCell.Value & "" & " = " & y
ActiveCell.Offset(1, 0).Select: Range(ActiveCell, ActiveCell.Offset(0, 10)).Select
ActiveCell.Value = "I?e O=" & x & " Y=" & y
ActiveCell.Offset(2, 0).Select
End Sub


