Attribute VB_Name = "Module1"
Sub UniqueTickerID()
Dim RowNum As Double
Dim Tickers As Range
RowNum = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To RowNum
Cells(i, 9).Value = Cells(i, 1).Value
Next i
Set Tickers = Range("I:I")
Tickers.RemoveDuplicates Columns:=1
End Sub
Sub StockWatcher()
Dim StockVol As Long
Dim RowNum As Double
Dim TickerNum As Double
'Number of ticker instances
RowNum = Cells(Rows.Count, 1).End(xlUp).Row
'Number of UNIQUE ticker instances
TickerNum = Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To RowNum
For j = 2 To TickerNum
StockVol = 0
If Cells(i, 1).Value = Cells(j, 9).Value Then
If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
StockVol = Cells(i, 7).Value + Cells(i + 1, 7).Value
Cells(j, 12).Value = StockVol
End If
End If
Next j
Next i
End Sub
Sub YearlyChange()
Dim StockOpen As Double
Dim StockClose As Double
Dim RowNum As Double
Dim TickerNum As Double
Dim YearlyChange As Double
Dim PercentChange As Double
RowNum = Cells(Rows.Count, 1).End(xlUp).Row
TickerNum = Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To RowNum
For j = 2 To TickerNum
StockOpen = 0
StockClose = 0
If Cells(i, 1).Value = Cells(j, 9).Value Then
If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
StockOpen = Cells(i, 3).Value + Cells(i + 1, 3).Value
StockClose = Cells(i, 6).Value + Cells(i + 1, 6).Value
YearlyChange = StockClose - StockOpen
PercentChange = YearlyChange / StockOpen
Cells(i, 10).Value = YearlyChange
Cells(i, 11).Value = PercentChange
End If
End If
Next j
Next i
End Sub
