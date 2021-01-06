Attribute VB_Name = "Module1"
Sub stockLoop()

' Insert headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Set initial variables
Dim tickerSymbol As String
Dim high_price As Double
Dim low_price As Double
Dim open_price As Double
Dim cloes_price As Double
Dim volume As Double
Dim yearlyChange As Double
Dim percentChange As Double
tickerSymbol = " "

' Track ticker symbol in summary table
Dim Summary_Ticker_Row As Double
Summary_Ticker_Row = 2

' Indicate number of existing data rows
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Indicate open and closing prices
open_price = Cells(2, 3).Value

' Loop through all ticker symbols
For i = 2 To lastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        tickerSymbol = Cells(i, 1).Value
        
        ' Calculate yearly change
        close_price = Cells(i, 6).Value
        yearlyChange = close_price - open_price
       
        'Calculate percentage change
             If open_price <> 0 Then
                percentChange = (yearlyChange / open_price)
            Else
                percentChange = 0
            End If
            
        open_price = Cells(i + 1, 3).Value

        ' Indicate volume
        volume = volume + Cells(i, 7).Value
    
        ' Print ticker symbol
        Range("I" & Summary_Ticker_Row).Value = tickerSymbol
    
        ' Print volume
        Range("L" & Summary_Ticker_Row).Value = volume
        
        ' Print yearly change and condiitonal color
        Range("J" & Summary_Ticker_Row).Value = yearlyChange
            If yearlyChange > 0 Then
                Range("J" & Summary_Ticker_Row).Interior.ColorIndex = 10
            Else
                Range("J" & Summary_Ticker_Row).Interior.ColorIndex = 3
            End If
            
        ' Print percentage change
        Range("K" & Summary_Ticker_Row).Value = Format(percentChange, "Percent")
        
        ' Add one to summary table row
        Summary_Ticker_Row = Summary_Ticker_Row + 1
        
        ' Reset volume
        volume = 0
    
    ' If data in next row is the same
    
    Else

     volume = volume + Cells(i, 7).Value

    End If

Next i


End Sub

