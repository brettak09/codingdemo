sub sheetloop()
Dim ws as worksheet
    For Each ws in Worksheets 
    ws.activate
        debug.print ws.name
    Next

End Sub