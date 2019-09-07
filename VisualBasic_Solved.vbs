Sub Homework_Two()

Dim Ticker As String
Dim op As Double
Dim cl As Double
Dim opcl As Double
Dim popcl As Double
Dim Summary As Long
Dim Total As Double
Dim LastRow As Long
Dim newt As String
Dim newt2 As String
Dim newt3 As String
Dim stock As Double
Dim gi As Double
Dim go As Double
Dim LastRow2 As Long

totalws = ActiveWorkbook.Worksheets.Count

For x = 1 To totalws

    hoja = Worksheets.Item(x).Name
    Worksheets(hoja).Select

Ticker = Cells(2, 1).Value
op = Cells(2, 6).Value
cl = 0
Summary = 2
Total = 0
LastRow = Range("B" & Rows.Count).End(xlUp).Row

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        cl = Cells(i, 6).Value
    
    If op = 0 Then
        popcl = 0
    Else
        popcl = ((cl - op) / op)
    End If
    
        opcl = cl - op
        Total = Total + Cells(i, 7).Value

        Range("K" & Summary).Value = Ticker
        Range("L" & Summary).Value = Format(opcl, "General Number")
        Range("M" & Summary).Value = FormatPercent(popcl, vbTrue)
        Range("N" & Summary).Value = Total

        Summary = Summary + 1
    
        Total = 0
        op = Cells(i + 1, 6).Value
        Ticker = Cells(i + 1, 1).Value
    
    Else
        Total = Total + Cells(i, 7).Value!

    If Range("L" & Summary).Value < 0 Then
        Range("L" & Summary).Interior.ColorIndex = 3
    Else
        Range("L" & Summary).Interior.ColorIndex = 4
    End If
    
    End If
    
Next i

gi = 0
go = 0
stock = 0
newt = 0
newt2 = 0
newt3 = 0
LastRow2 = Range("M" & Rows.Count).End(xlUp).Row
gi = Cells(2, 13).Value
go = Cells(2, 13).Value
stock = Cells(2, 14).Value

For j = 2 To LastRow2

If Cells(j + 1, 13).Value > gi Then
    gi = Cells(j + 1, 13).Value
    newt = Cells(j + 1, 11).Value
End If

If Cells(j + 1, 13).Value < go Then
    go = Cells(j + 1, 13).Value
    newt2 = Cells(j + 1, 11).Value
End If

If Cells(j + 1, 14).Value > stock Then
    stock = Cells(j + 1, 14).Value
    newt3 = Cells(j + 1, 11).Value
End If

Next j

Range("R2").Value = newt
Range("S2").Value = FormatPercent(gi, vbTrue)
Range("R3").Value = newt2
Range("S3").Value = FormatPercent(go, vbTrue)
Range("R4").Value = newt3
Range("S4").Value = stock

Next x

End Sub
