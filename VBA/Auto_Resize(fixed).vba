Sub AutoFitEverything()
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ActiveSheet
    Set rng = ws.UsedRange

    Application.ScreenUpdating = False

    ' 1. Unmerge all cells (AutoFit cannot work otherwise)
    rng.UnMerge

    ' 2. Disable Wrap Text (critical for column autofit)
    rng.WrapText = False

    ' 3. Force formulas to recalc as values for sizing
    rng.Value = rng.Value

    ' 4. AutoFit columns first
    rng.Columns.AutoFit

    ' 5. Then AutoFit rows
    rng.Rows.AutoFit

    Application.ScreenUpdating = True
End Sub
