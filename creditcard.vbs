sub CardSummary()
    'iterate across all rows with data
    'while the card is unchanged, keep incrimenting a card subtotal value
    'when the card changes, write the value to a summary row or table
    '  or just always write the subtotal to the column D
end sub

Sub
Sub card_summary()

    ' figure out what the end is, xlUp, xlLeft, xlRight
    rowEnd = Cells(2, 1).End(xlDown).Row
Subtotal = 0

    For Row = 2 To rowEnd
        current_card = Cells(Row, 1).Value
        next_stock = Cells(Row + 1, 1).Value

        Subtotal = Subtotal + Cells(Row, 3).Value
        Cells(Row, 4).Value = Subtotal

        If current_card <> next_card Then
            Subtotal = 0
        End If

    Next Row

End Sub

Sub cardSummary()

   ' figure out what the end is, xlUp, xlLeft, xlRight
   rowEnd = Cells(2, 1).End(xlDown).Row
   subtotal = 0
   rowIndex_table = 2

   For row = 2 to rowEnd
       current_card = Cells(row, 1).Value
       next_card = Cells(row + 1, 1).Value
       subtotal = subtotal + Cells(row,3).Value
       Cells(row,4).Value = subtotal

       If current_card <> next_card Then
           Cells(rowIndex_table, 7).Value = current_card
           Cells(rowIndex_table, 8).Value = subtotal
           rowIndex_table = rowIndex_table + 1
           subtotal = 0
       End If
   Next row
End Sub