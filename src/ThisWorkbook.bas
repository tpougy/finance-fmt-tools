Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Const PROC As String = "Workbook_BeforeClose"

    If ThisWorkbook.Saved = False Then
        On Error GoTo ErrHandler
        Application.EnableEvents = False
        ThisWorkbook.Save
        Application.EnableEvents = True
        Log PROC & ": .xlam salvo como fallback no fechamento"
    End If
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Log PROC & ": ERRO no save de fechamento — configurações podem não ter persistido"
    HandleError PROC, Err
End Sub