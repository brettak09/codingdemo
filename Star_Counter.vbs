Sub nestedForloop()

For rowIndex = 2 to 51
      total = 0
   For colIndex = 4 to 8
   
       If cells(rowIndex, colIndex).Value = "Full-Star" Then
           total = total + 1
       Else
           total = total
       End If
   Next colIndex

   cells(rowIndex,9).Value = total

Next rowIndex
sub nestedForloop