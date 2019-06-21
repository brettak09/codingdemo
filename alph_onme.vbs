Sub alpha_onMe()
' set an initial variable for hodling the ticker name
Dim ticker_name as String
' Set an Initial variable for holding the total per ticker Brand
    Dim ticker_total As Double
        ticker_total = 0
'Keep track of the location for each ticker name in teh summary table
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
'Loop through all ticker vol
    For vol_loop = 2 to 70926
'Check if we are still within the same ticker, if we are not...
    If Cells(vol_loop + 1, 1).Value <> Cells(vol_loop, 1).Value Then
'set the ticker name
    ticker_name = Cells(vol_loop, 1).Value
'Add to the ticker total
    ticker_total = ticker_total + Cells(vol_loop, 7).Value
'Print the ticker name in the summary table
    Range("I" & Summary_Table_Row).Value = ticker_name
'Print the ticker Amount to tthe Summary table
    Range("J" & Summary_Table_Row).Value = ticker_total
'add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
'reset the ticker total
    ticker_total = 0
'If the cell immedialtly following a row is the same brand...
    Else
'add to the ticker total
    ticker_total = ticker_total + Cells(vol_loop, 7).Value

'message box the unique ticker name
            
     End if
    
    Next vol_loop

End Sub
