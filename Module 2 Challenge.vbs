Attribute VB_Name = "Module1"
Sub RunThroughWorkbook()

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

Stock_Data
Plus_Zero
Yearly_Change
Add_Functionality

Next ws


End Sub


Sub Stock_Data()

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"

Dim ticker As String
Dim ticker2 As String
Dim i As Long
Dim lastrow As Long
Dim counter As Integer

counter = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
ticker = Range("A" & i).Value
ticker2 = Range("A" & i + 1).Value

If ticker <> ticker2 Then
Range("I" & counter) = ticker

counter = counter + 1

End If

Next i

End Sub


Sub Plus_Zero()

Dim lastrow As Long
Dim i As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
Range("B" & i).Value = Range("B" & i).Value + 0

Next i


End Sub

Sub Yearly_Change()

Dim MaxDate As Long
Dim MinDate As Long
Dim i As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentageChange As Double
Dim TotalStockVolume As LongLong

lastrow = Cells(Rows.Count, 9).End(xlUp).Row

MaxDate = WorksheetFunction.Max(Range("B:B"))
MinDate = WorksheetFunction.Min(Range("B:B"))

For i = 2 To lastrow

OpenPrice = WorksheetFunction.SumIfs(Range("C:C"), Range("A:A"), Range("I" & i), Range("B:B"), MinDate)

ClosePrice = WorksheetFunction.SumIfs(Range("F:F"), Range("A:A"), Range("I" & i), Range("B:B"), MaxDate)

YearlyChange = ClosePrice - OpenPrice
Range("J" & i).Value = YearlyChange

PercentageChange = YearlyChange / OpenPrice
Range("K" & i).Value = PercentageChange

TotalStockVolume = WorksheetFunction.SumIfs(Range("G:G"), Range("A:A"), Range("I" & i))
Range("L" & i).Value = TotalStockVolume

Next i

End Sub

Sub Add_Functionality()

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestTotalVolume As LongLong
Dim GreatestPercentIncreaseTicker As Integer
Dim GreatestPercentDecreaseTicker As Integer
Dim GreatestTotalVolumeTicker As Integer

GreatestPercentIncrease = WorksheetFunction.Max(Range("K:K"))
GreatestPercentDecrease = WorksheetFunction.Min(Range("K:K"))
GreatestTotalVolume = WorksheetFunction.Max(Range("L:L"))

Range("Q2").Value = GreatestPercentIncrease
Range("Q3").Value = GreatestPercentDecrease
Range("Q4").Value = GreatestTotalVolume

GreatestPercentIncreaseTicker = WorksheetFunction.Match(Range("Q2").Value, Range("K:K"), 0)
Range("P2").Value = Range("I" & GreatestPercentIncreaseTicker)

GreatestPercentDecreaseTicker = WorksheetFunction.Match(Range("Q3").Value, Range("K:K"), 0)
Range("P3").Value = Range("I" & GreatestPercentDecreaseTicker)

GreatestTotalVolumeTicker = WorksheetFunction.Match(Range("Q4").Value, Range("L:L"), 0)
Range("P4").Value = Range("I" & GreatestTotalVolumeTicker)


End Sub






