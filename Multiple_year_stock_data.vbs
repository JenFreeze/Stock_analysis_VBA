Attribute VB_Name = "Module1"
Sub StockLoop()

For Each ws In Worksheets

'Add Labels to Each Column
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Determine Last Row of Main Dataset
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Declare Variables Needed
Dim i As Long 'loop variable
Dim SumRow As Double 'summarized table row
SumRow = 2

Dim Ticker As String
Dim OpenValue As Double
Dim CloseValue As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Double

'Loop to Get List of Unique Tickers, Yearly Change, Percent Change, and Total Volume
For i = 2 To LastRow
    Ticker = ws.Cells(i, 1).Value
    If Ticker <> ws.Cells(i - 1, 1).Value Then
        OpenValue = ws.Cells(i, 3)
        ws.Cells(SumRow, 9) = Ticker
        Volume = Volume + ws.Cells(i, 7)
    ElseIf Ticker <> ws.Cells(i + 1, 1).Value Then
        CloseValue = ws.Cells(i, 6)
        YearlyChange = CloseValue - OpenValue
        ws.Cells(SumRow, 10) = YearlyChange
        PercentChange = YearlyChange / OpenValue
        ws.Cells(SumRow, 11) = PercentChange
        Volume = Volume + ws.Cells(i, 7)
        ws.Cells(SumRow, 12) = Volume
        Volume = 0 'reset Volume to 0 after last row with certain ticker
        SumRow = SumRow + 1
    Else
        Volume = Volume + ws.Cells(i, 7)
    End If

Next i

'Format Yearly Change and Percent Change Columns
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("J:J").NumberFormat = "0.00"

'Calculate Greatest Increase, Decrease, and Total Volume
'Add Labels to Table
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest & Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Declare and Set Variable Values
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

Dim LastRowSummary As Long
LastRowSummary = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Calculate Each Amount and Add Conditional Formatting to Yearly Change Column
For SumRow = 2 To LastRowSummary

    If ws.Cells(SumRow, 10).Value < 0 Then
        ws.Cells(SumRow, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(SumRow, 10).Interior.ColorIndex = 4
    End If
    
    If ws.Cells(SumRow, 11).Value > GreatestIncrease Then
        ws.Range("P2").Value = ws.Cells(SumRow, 9).Value
        ws.Range("Q2").Value = ws.Cells(SumRow, 11).Value
        GreatestIncrease = ws.Cells(SumRow, 11).Value
    End If
    
    If ws.Cells(SumRow, 11).Value < GreatestDecrease Then
        ws.Range("P3").Value = ws.Cells(SumRow, 9).Value
        ws.Range("Q3").Value = ws.Cells(SumRow, 11).Value
        GreatestDecrease = ws.Cells(SumRow, 11).Value
    End If
    
    If ws.Cells(SumRow, 12).Value > GreatestVolume Then
        ws.Range("P4").Value = ws.Cells(SumRow, 9)
        ws.Range("Q4").Value = ws.Cells(SumRow, 12)
        GreatestVolume = ws.Cells(SumRow, 12).Value
    End If
    
Next SumRow

ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

End Sub

