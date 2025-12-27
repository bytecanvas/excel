Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Only resize the changed cells
    Target.WrapText = False
    Target.EntireColumn.AutoFit
    Target.EntireRow.AutoFit

ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
