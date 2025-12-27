Sub AutoFitSafeButStrong()
    With ActiveSheet.UsedRange
        .WrapText = False
        .Columns.AutoFit
        .Rows.AutoFit
    End With
End Sub
