Sub stockInfo()
' Loop through all the worksheets '
Dim WS As Worksheet
'  Call active worksheet
    For Each WS In ActiveWorkbook.Worksheets  
       WS.Activate
    ' Determine the Last row
    last_row = WS.Cells(Rows.Count, 1).End(xlUp).row
    ' Declare variables
    Dim open_price As Double
    Dim closing_price As Double
    Dim annual_change As Double
    Dim ticker As String
    Dim percentage_change As Double
    
    ' setting the volume variable to 0
    Dim volume As Double
    volume = 0
    ' setting column to 1
    Dim column As Integer
    column = 1
    ' setting row to 2
    Dim row As Double
    row = 2
    Dim i As Long
    ' headings for summary columns row 1, columns I, J,K, and L)
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
        
    'Set Opening Price
    open_price = Cells(2, column + 2).Value
' BEGIN LOOP
                
        For i = 2 To last_row
         ' Pass through and determine if we are within the same ticker, ELSE
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                ' write the name of the ticker out to the row and column
                ticker = Cells(i, column).Value
                Cells(row, column + 8).Value = ticker
                ' Set closing price
                closing_price = Cells(i, column + 5).Value
                ' Get the annual change by closing price - open price
                annual_change = closing_price - open_price
                Cells(row, column + 9).Value = annual_change
                ' Create percentage change loop.
                ' If the open and close price are 0
                ' then the percentage change is 0.
                ' Else If the open price is 0 and the closing price <> 0
                ' then the percentage chane is 1.
                If (open_price = 0 And closing_price = 0) Then
                    percentage_change = 0
                ElseIf (open_price = 0 And closing_price <> 0) Then
                    percentage_change = 1
                ' Else the percentage change is the annual change  divided by the open price
                Else
                    percentage_change = annual_change / open_price
                    Cells(row, column + 10).Value = percentage_change
                    'ADDED BACK TO GET %
                    Cells(row, column + 10).NumberFormat = "0.00%"  '  hard coding the number format
                End If
                ' get the total volume by adding cells together
                volume = volume + Cells(i, column + 6).Value
                Cells(row, column + 11).Value = volume
                ' write to the summary
                row = row + 1
                ' Reset the open price and reset the volume variable back to 0
                open_price = Cells(i + 1, column + 2)
                volume = 0
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
        Next i
        
' Determine the annual change of last row per WS
        annual_change_last_row = WS.Cells(Rows.Count, column + 8).End(xlUp).row
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Set red and green colors for cells
        For j = 2 To annual_change_last_row
            If (Cells(j, column + 9).Value > 0 Or Cells(j, column + 9).Value = 0) Then
                Cells(j, column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, column + 9).Value < 0 Then
                Cells(j, column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Write Greatest % Increase, % Decrease, and Total volume
        Cells(2, column + 14).Value = "Greatest % Increase"
        Cells(3, column + 14).Value = "Greatest % Decrease"
        Cells(4, column + 14).Value = "Greatest Total volume"
        Cells(1, column + 15).Value = "Ticker"
        Cells(1, column + 16).Value = "Value"
        ' Search through each of the rows to find the greatest value and it's respective ticker
        
        For Z = 2 To annual_change_last_row
            If Cells(Z, column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & annual_change_last_row)) Then
                Cells(2, column + 15).Value = Cells(Z, column + 8).Value
                Cells(2, column + 16).Value = Cells(Z, column + 10).Value
                Cells(2, column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & annual_change_last_row)) Then
                Cells(3, column + 15).Value = Cells(Z, column + 8).Value
                Cells(3, column + 16).Value = Cells(Z, column + 10).Value
                Cells(3, column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & annual_change_last_row)) Then
                Cells(4, column + 15).Value = Cells(Z, column + 8).Value
                Cells(4, column + 16).Value = Cells(Z, column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub







