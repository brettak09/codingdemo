Sub stock_value()

   ' figure out what the end is, xlUp, xlLeft, xlRight
   rowEnd = Cells(1, 1).End(xlDown).Row
   Subtotal = 0
   rowIndex_table = 2

   For Row = 2 To rowEnd
       Ticker = Cells(Row, 1).Value
       Next_Ticker = Cells(Row + 1, 1).Value
       Subtotal = Subtotal + Cells(Row, 7).Value
       

    If Ticker <> Next_Ticker Then
           Cells(rowIndex_table, 10).Value = Ticker
           Cells(rowIndex_table, 11).Value = Subtotal
           rowIndex_table = rowIndex_table + 1
           Subtotal = 0
    End If
   Next Row
End Sub