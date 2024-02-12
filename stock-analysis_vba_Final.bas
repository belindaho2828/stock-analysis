Attribute VB_Name = "Module1"
Sub Ticker()

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim StockVolume As LongLong
Dim SumTableRow As Integer
Dim EndRow As Long
Dim i As Long
Dim YEDelta As Double
Dim TickerOpen As Double
Dim TickerClose As Double
Dim HighestDelta As Double
Dim TickerHighest As String
Dim TickerHighestVolume As String
Dim TickerLowest As String
Dim LowestDelta As Double
Dim HighestVolume As LongLong
Dim ChangeEndRow As Integer
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Setting headers for 2nd summary table (output table)
Cells(1, "I") = "Ticker"
Cells(1, "J") = "Yearly Change"
Cells(1, "K") = "Percent Change"
Cells(1, "L") = "Total Stock Volume"

'Creating headers for the 2nd summary table
Cells(2, "O").Value = "Greatest % Increase"
Cells(3, "O").Value = "Greatest % Decrease"
Cells(4, "O").Value = "Greatest Total Value"
Cells(1, "P").Value = "Ticker"
Cells(1, "Q").Value = "Value"

'Setting values at 0 so that the first loop will compare to this first
HighestDelta = 0
LowestDelta = 0
HighestVolume = 0

SumTableRow = 2

EndRow = Cells(Rows.Count, "A").End(xlUp).Row

'For Loop

For i = 2 To EndRow
    
    'Store the Ticker at each loop
    Ticker = Cells(i, "A").Value
    
    'Store TickerClose at each loop
    TickerClose = Cells(i, "F").Value
    
    'Firstrow logic: If the current ticker symbol is different from the ticker symbol below it
    If Cells(i, "A").Value <> Cells(i - 1, "A").Value Then
        
        'Store Ticker Open for the first row of each Ticker
        TickerOpen = Cells(i, "C").Value
        
        'Test TickerOpen price:
        'Range("O" & SumTableRow).Value = TickerOpen
        
        'Print the ticker in SumTableRow
        Range("I" & SumTableRow).Value = Ticker
        
        'Add one to the summary Table Row (for new ticker)
        SumTableRow = 1 + SumTableRow
    
        'Store the Stock Volume at each loop
        StockVolume = Cells(i, "G").Value
    
    Else
    
        'Add to the Stock Volume
        StockVolume = StockVolume + Cells(i, "G").Value
    
        'Last row logic
        If Cells(i, "A").Value <> Cells(i + 1, "A").Value Then
        
            'Set delta of the Close price less than ticker's open price at beg. of year
            YEDelta = TickerClose - TickerOpen
            
            'If stock volume is greater than highest volume (finding max), record the stockvolume in HighestVolume variable. This gets overwritten at every higher comparison.
            If StockVolume > HighestVolume Then
                HighestVolume = StockVolume
                TickerHighestVolume = Cells(i, "A").Value
             End If
            
            'Test Close price:
            'Range("N" & SumTableRow - 1).Value = TickerClose
            
            'Print the Yearly Change
            Range("J" & SumTableRow - 1).Value = YEDelta
            
            'Set % change of Close price to open price
            PercentChange = YEDelta / TickerOpen
            
            'If % Change is greater than highest Delta (finding max), record the % Change in HighestDelta variable. This gets overwritten at every higher comparison.
            If PercentChange > HighestDelta Then
                HighestDelta = PercentChange
                HighestTicker = Cells(i, "A").Value
            End If
            
            'If % Change is lower than lowest Delta (finding min), record the % Change in LowestDelta variable. This gets overwritten at every lower comparison.
            If PercentChange < LowestDelta Then
                LowestDelta = PercentChange
                LowestTicker = Cells(i, "A").Value
            End If
            
             'Testing for negative Yearly Change and setting to red if true for all worskheets
            If YEDelta < 0 Then
                Range("J" & SumTableRow - 1).Interior.Color = RGB(255, 0, 0)
            Else
                'Otherwise, set the cell to green for all worksheets
                Range("J" & SumTableRow - 1).Interior.Color = RGB(0, 255, 0)
        
            End If
            
            'Testing for negative % Change and setting to red if true for all worksheets
            If PercentChange < 0 Then
                Range("K" & SumTableRow - 1).Interior.Color = RGB(255, 0, 0)
            Else
                'Otherwise, set the cell to green for all worksheets
                Range("K" & SumTableRow - 1).Interior.Color = RGB(0, 255, 0)
            
            End If
            
            'Print the T Change in SumTableRow and set formatting to percentage
            Range("K" & SumTableRow - 1).Value = PercentChange
            Range("K" & SumTableRow - 1).NumberFormat = "0.00%"
            
            'Print the StockVolume in SumTableRow (current ticker)
            Range("L" & SumTableRow - 1).Value = StockVolume
    
    
        End If 'End If of last row logic
    
    End If 'End If change at Ticker logic

Next i

'Print HighestDelta amount in the 3rd summary table (output table) and the associated ticker + set number format
Cells(2, "Q").Value = HighestDelta
Cells(2, "Q").NumberFormat = "0.00%"
Cells(2, "P").Value = HighestTicker

'Print LowestDelta Ticker and value that was stored above
Cells(3, "Q").Value = LowestDelta
Cells(3, "Q").NumberFormat = "0.00%"
Cells(3, "P").Value = LowestTicker

'Print HighestDelta Ticker and value that was stored above
Cells(4, "Q").Value = HighestVolume
Cells(4, "P").Value = TickerHighestVolume

Next ws

End Sub
