sub chcken_nuggets()
    For chicken_nuggests = 1 To 16
        cells(chicken_nuggests, 3).Value = "Chicken Nuggets"
            if cells(16, 2).Value > 10 then
                cells(chicken_nuggests, 1).Value = "I will Eat"
            Else 
                cells(chicken_nuggests, 1).Value = "I won't Eat"
            End if
        
    Next chicken_nuggests

End sub