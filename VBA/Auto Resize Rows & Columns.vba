Sub AutoResizeCells()
    Dim ws As Worksheet
    
    ' Set the worksheet to the active sheet (you can change this if needed)
    Set ws = ActiveSheet
    
    ' Automatically adjust the width of columns to fit their content
    ws.Columns.AutoFit
    
    ' Automatically adjust the height of rows to fit their content
    ws.Rows.AutoFit
End Sub
