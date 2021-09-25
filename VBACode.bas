Attribute VB_Name = "Module1"
Sub StockCharts():

' variables
Dim ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double

' loop over each worksheet in the workbook
For Each ws In Worksheets

    ' Make the worksheet active.
    ws.Activate

    ' Find the last row of each worksheet
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header columns for each worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Initialize variables for each worksheet.
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' Skipping the header row, loop through the list of tickers.
    For i = 2 To lastRowState

        ' Get the value of the ticker symbol we are currently calculating for.
        ticker = Cells(i, 1).Value
        
        ' Get the start of the year opening price for the ticker.
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Add up the total stock volume values for a ticker.
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' Run this if we get to a different ticker in the list.
        If Cells(i + 1, 1).Value <> ticker Then
            ' Increment the number of tickers when we get to a different ticker in the list.
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            ' Get the end of the year closing price for ticker
            closing_price = Cells(i, 6)
            
            ' Get yearly change value
            yearly_change = closing_price - opening_price
            
            ' Add yearly change value to the appropriate cell in each worksheet.
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' If yearly change value is greater than 0, shade cell green.
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ' If yearly change value is less than 0, shade cell red.
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            ' If yearly change value is 0, shade cell yellow.
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Calculate percent change value for ticker.
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            ' Format the percent_change value as a percent.
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            ' Set opening price back to 0 when we get to a different ticker in the list.
            opening_price = 0
            
            ' Add total stock volume value to the appropriate cell in each worksheet.
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Set total stock volume back to 0 when we get to a different ticker in the list.
            total_stock_volume = 0
        End If
        
    Next i
    
Next ws


End Sub

'reference code from https://github.com/philipstubbs13/vba-challenge

