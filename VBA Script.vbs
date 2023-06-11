Sub Stocks()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets

ws.Activate

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"


Dim Ticker As String
Dim i As Double

Dim Lastrow As Double
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim Row As Double
Row = 2

Dim Volume As Double
Volume = 0

Dim Open1 As Variant
Open1 = Empty
Dim Close1 As Double

Dim Yearly_Change As Double
Dim Percentage As Double


For i = 2 To Lastrow

Ticker = Cells(i, 1).Value
Volume = Volume + Cells(i, 7).Value


If Open1 = Empty Then

Open1 = Cells(i, 3).Value

End If


Close1 = Cells(i, 6).Value


If Ticker <> Cells(i + 1, 1).Value Then

Range("I" & Row).Value = Ticker
Range("L" & Row).Value = Volume
Range("J" & Row).Value = Close1 - Open1
Range("K" & Row).Value = (Close1 - Open1) / Open1
Range("K" & Row).NumberFormat = "0.00%"


Volume = 0
Open1 = Empty

Row = Row + 1


End If


Next i



Dim j As Double

For j = 2 To Lastrow

Yearly_Change = Cells(j, 10).Value
Percentage = Cells(j, 11).Value


If Yearly_Change > 0 Then
Cells(j, 10).Interior.ColorIndex = 4

Else

Cells(j, 10).Interior.ColorIndex = 3

End If

If Percentage > 0 Then

Cells(j, 11).Interior.ColorIndex = 4

Else

Cells(j, 11).Interior.ColorIndex = 3

End If

Next j

'BONUS

Range("P1").Value = "Ticker"
Range("Q1").Value = "Volume"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest total volume"


Dim Rng As Range
Set Rng = Range("K1:K" & Lastrow)

Dim Max As Double
Dim MaxRow As Double
Dim MaxTicker As String


Max = Application.WorksheetFunction.Max(Rng)
MaxRow = Application.WorksheetFunction.Match(Max, Rng, 0)
MaxTicker = Range("I" & MaxRow).Value

Range("P2").Value = MaxTicker
Range("Q2").Value = Max
Range("Q2").NumberFormat = "0.00%"

Dim Min As Double
Dim MinRow As Double
Dim MinTicker As String

Min = Application.WorksheetFunction.Min(Rng)
MinRow = Application.WorksheetFunction.Match(Min, Rng, 0)
MinTicker = Range("I" & MinRow).Value

Range("P3").Value = MinTicker
Range("Q3").Value = Min
Range("Q3").NumberFormat = "0.00%"


Dim Rng2 As Range
Set Rng2 = Range("L1:L" & Lastrow)

Dim Max2 As Double
Dim Max2Row As Double
Dim Max2Ticker As String


Max2 = Application.WorksheetFunction.Max(Rng2)
Max2Row = Application.WorksheetFunction.Match(Max2, Rng2, 0)
Max2Ticker = Range("I" & Max2Row).Value

Range("P4").Value = Max2Ticker
Range("Q4").Value = Max2




Next ws


End Sub





