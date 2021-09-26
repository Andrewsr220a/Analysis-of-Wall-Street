Sub Stock_market()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets


'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Defining Variables

'Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
TickerVolume = 0

'Open and close price variables
Dim open_price As Double
open_price = 0

Dim close_price As Double
close_price = 0

'Yearly change
Dim Pricechange As Double
Pricechange = 0

'Total stock_volume
Dim stock_volume As Long
StockVolume = 0

'Percent Equation for bonus
Dim PercentChange As Double
PercentChange = 0

DispRow = 2

            'Set initial and last row for worksheet
            Dim RowCount As Long
            Dim i As Long
            Dim j As Integer

            RowCount = Cells(Rows.Count, 1).End(xlUp).Row

            'loop all the rows '
            For i = 2 To RowCount
            If open_price = 0 Then
            open_price = ws.Cells(i, 3).Value
            
            End If
            
            'When the tickers are not equal -it is at the last row for the ticker
            
            'Ticker Column
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
             
             DispRow = DispRow + 1
            'Reset variables
                open_price = 0
                open_price = ws.Cells(i + 1, 1).Value
                
            'When the tickers are equal
            
            'Fixing the open price equal zero problem
            
            ElseIf open_price <> 0 Then
            Pricechange = close_price - open_price
            PriceChangePercent = (Pricechange / open_price) * 100
 

                'Year change column
                ClosePrice = ws.Cells(i, 6).Value
                OpenPrice = ws.Cells(i, 3).Value
            
                YearChange = (ClosePrice - OpenPrice)
                Cells(2, 10).Value = YearChange
                
                'Percent Change column
                PercentChange = (PercentChange / OpenPrice) * 100
                ws.Cells(2, 11).Value = YearChange
                
                

                
            End If

Next i

Next ws

End Sub
