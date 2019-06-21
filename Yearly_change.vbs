Sub yearly_change_by_date()
'identify the ticker name, first dates and last date'
    Dim ticker_name as String
    Dim Firstdate as Double
    Dim Lastdate as Double
'Initial variable for hodling yearly change
    Dim yearly_change as long
        yearly_change = 0
'keep track of the location for yearly_change in sum table 
    Dim Summary_Table_Row As Double
            Summary_Table_Row = 2
'value of first and last date
    Firstdate = 20150101
    Lastdate = 20151231
'loop through cells of ticker
    for ticker_loop = 1 to 262
'Check if we are still within the same ticker, if we are not...
    if Cells(ticker_loop + 1, 2).Value <> Cells(ticker_loop, 2).Value Then
'''set a name for the yearly_change_calculation ''''
'find the first and last date value
'''''Issues here****
    Cells(ticker_loop, 6).Value = Firstdate 
    Cells(ticker_loop, 6).Value = Lastdate   
'Calculate the value of the close price 
'for firstdate and subtract it by the close price for the lastdate
    yearly_change + (Cells(ticker_loop, 6).Value = Firstdate) * Cells(ticker_loop, 7).Value = Lastdate
'Print the yealry_change Amount to the Summary table
    Range("J" & Summary_Table_Row).Value = yearly_change

    End If
    next ticker_loop



End Sub