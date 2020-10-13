Sub Calculate_Stock_Stats():

' variable to keep track of current ticker symbol
Dim ticker As String

' variable to keep track of number of tickers for each worksheet
Dim number_tickers As Integer

' variable to keep track of the last row in each worksheet.
Dim lastRowState As Long

' variable to keep track of opening price for specific year
Dim opening_price As Double

' variable to keep track of closing price for specific year
Dim closing_price As Double

' variable to keep track of yearly change
Dim yearly_change As Double

' variable to keep track of percent change
Dim percent_change As Double

' variable to keep track of total stock volume
Dim total_stock_volume As Double

' variable to keep track of greatest percent increase value for specific year.
Dim greatest_percent_increase As Double

' variable to keep track of the ticker that has the greatest percent increase.
Dim greatest_percent_increase_ticker As String

' varible to keep track of the greatest percent decrease value for specific year.
Dim greatest_percent_decrease As Double

' variable to keep track of the ticker that has the greatest percent decrease.
Dim greatest_percent_decrease_ticker As String

' variable to keep track of the greatest stock volume value for specific year.
Dim greatest_stock_volume As Double

' variable to keep track of the ticker that has the greatest stock volume.
Dim greatest_stock_volume_ticker As String

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
            
            
            ' Uncomment the following for color shading of percent change column.
            ' If percent change value is greater than 0, shade cell green.
            ' If percent_change > 0 Then
                ' Cells(number_tickers + 1, 11).Interior.ColorIndex = 4
            ' If percent change value is less than 0, shade cell red.
            ' ElseIf percent_change < 0 Then
                ' Cells(number_tickers + 1, 11).Interior.ColorIndex = 3
            ' If percent change value is 0, shade cell yellow.
            ' Else
                ' Cells(number_tickers + 1, 11).Interior.ColorIndex = 6
            ' End If
            
            
            ' Set opening price back to 0 when we get to a different ticker in the list.
            opening_price = 0
            
            ' Add total stock volume value to the appropriate cell in each worksheet.
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Set total stock volume back to 0 when we get to a different ticker in the list.
            total_stock_volume = 0
        End If
        
    Next i
    
    ' Add section to display greatest percent increase, greatest percent decrease, and greatest total volume for each year.
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Get the last row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables and set values of variables initially to the first row in the list.
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    ' skipping the header row, loop through the list of tickers.
    For i = 2 To lastRowState
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest percent decrease.
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest stock volume.
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Add the values for greatest percent increase, decrease, and stock volume to each worksheet.
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub

'Create a script that will loop through all the stocks for one year for each run and take the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.

    'You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub stocks2()

'loop through each worksheet
Dim ws As Worksheet
For Each ws In Worksheets
    
    'set headers for summary table
    ws.Range("I1").Value = ("ticker")
    ws.Range("J1").Value = ("Yearly" & " " & "Change")
    ws.Range("k1").Value = ("Percent" & " " & "Change")
    ws.Range("l1").Value = ("Total" & " " & "Stock" & "Volume")

'------------------------------------------------------------------
'Calculate the tracker and the total volume, percent change and yearly change
'------------------------------------------------------------------

    ' Set variable for holding the ticker
    Dim ticker As String
    
    ' Set variable for the open values
    Dim open_value As Double
    Dim openvalue_ind As Double
    openvalue_ind = 2
    Dim yearlyvalue As Double
    Dim percentchange As Double
    
    'Set variable for the close value
    Dim close_value As Double
    
    ' Set variable for holding the total per ticker
    Dim tickertotal As Double
    tickertotal = 0
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Loop through all tickers
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        open_value = ws.Cells(openvalue_ind, 3).Value
        
    ' Check if still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Set the ticker
        ticker = ws.Cells(i, 1).Value
        
        ' Add to the ticker Total
        tickertotal = tickertotal + ws.Cells(i, 7).Value
        
        'Get the closevalue
        close_value = ws.Cells(i, 6).Value
        
        'calculate yearlychange
        yearlychange = close_value - open_value
        ws.Cells(i, 10).Value = yearlychange

            'formatting yearlychange
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        
        'calculate percentchange
        If open_value = 0 Then
            percentchange = 0
        Else
            percentchange = yearlychange / open_value
        End If
        
        'Print the values in the Summary Table
        ws.Range("K" & Summary_Table_Row).Value = percentchange
        'Format the percentage in the Summary Table
        ws.Range("K" & Summary_Table_Row) = Format(percentchange, "Percent")
        'Print the yearlychange in the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = yearlychange
        ' Print the ticker in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ' Print the tickertotal to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = tickertotal
    
        'format the yearlychange
        If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        End If
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the ticker Total
        tickertotal = 0
        
        ' Rest openvalue_ind to proper number
        openvalue_ind = (i + 1)
        
    ' If the cell immediately following a row is the same ticker
    Else

        ' Add to the ticker Total
        tickertotal = tickertotal + ws.Cells(i, 7).Value
        
    End If
    
  Next i

'------------------------------------------------------------------
'autofit formatting
'------------------------------------------------------------------
    'autofit all
    ws.Range("A:M").Columns.AutoFit

Next ws

End Sub
