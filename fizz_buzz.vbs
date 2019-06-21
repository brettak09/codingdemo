Sub fizz_buzz()
'create a loop to look at the values
For numbers = 2 To 100
'If the value in column 1 is a multiple of both 3 and 5, print "Fizzbuzz" in column 2.
    If Cells(numbers, 1).Value Mod 3 = 0 And Cells(numbers, 1).Value Mod 5 = 0 Then
            Cells(numbers, 2).Value = "Fizzbuzz"
        ElseIf (numbers Mod 3 = 0) Then
            Cells(numbers, 2).Value = "Fizz"
    ElseIf (numbers Mod 5 = 0) Then
            Cells(numbers, 2).Value = "buzz"

    End If
'If the value in column 1 is a multiple of just 3, print "Fizz" in column 2.

'If the value in column 1 is a multiple of just 5, print "Buzz" in column 2.

Next numbers


End Sub