Sub tickertotaler_moderate()

'''''Stock Volume is short for A.  Why?

'define everything
Dim ws As Worksheet
Dim ticker_name As String
Dim volume As Integer
Dim volume_total As Double
    volume_total = 0
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
'dim open_price As Boulean''''reveiew how this works'''

'this prevents my overflow error
On Error Resume Next

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup integers for loop
    Summary_Table_Row = 2

    'calculate yearly out
    year_open = ws.Cells(2, 3).Value

    'loop
        For I = 2 To ws.UsedRange.Rows.Count
        '''''get open price at the beginning of the year''''''
             If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            'find all the values
            ticker_name = ws.Cells(I, 1).Value
               
            'volume_total = ws.Cells(I, 7).Value

            
            year_close = ws.Cells(I, 6).Value

            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_open
            ws.Cells(i + 1, 3).Value = year_open
            'add in year_close
           ' ws.Cells(i + 1, 6).Value = year_close
            volume_total = volume_total + Cells(I, 7).Value
            'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker_name
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = volume_total
            Summary_Table_Row = Summary_Table_Row + 1

             volume_total = 0
             percent_change = 0
        Else 
            volume_total = volume_total + Cells(I, 7).Value
        
        End If

'finish loop
    Next I
    
ws.Columns("K").NumberFormat = "0.00%"


    'format columns colors
    Dim rg As Range
    Dim g As Double
    Dim c As Double
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g




'move to next worksheet
Next ws


End Sub