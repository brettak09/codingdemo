sub budget_checker()
    total = Range("F3").Value * (1+Range("H3").Value)
        Msgbox(total)
            budget = Range("C3").Value
                MsgBox(budget)
                    Range("L3").Value = total
                    If total < budget then
                        MsgBox("Under budget")
                    Elseif total > budget then
                        MsgBox("over budget")
                    End If
End Sub