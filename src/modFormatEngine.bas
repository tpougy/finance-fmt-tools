Option Explicit

' =============================================================================
' MOTOR CENTRAL DE FORMATAÇÃO
' Toda a lógica de formatação passa por aqui.
' Ribbon callbacks e chamadas diretas devem usar ApplyFormat().
' =============================================================================

' -- Tipo público para descrever um formato ------------------------------------
Public Type FormatDef
    key          As String
    DisplayName  As String
    NumberFmt    As String
    Category     As String      ' "numeric", "percent", "date", "text"
    Alignment    As XlHAlign    ' xlHAlignRight, xlHAlignLeft, xlHAlignGeneral
End Type


' =============================================================================
' PUBLIC API
' =============================================================================

' Ponto de entrada principal. Aceita qualquer Range (seleção ou não).
Public Sub ApplyFormat(ByVal rng As Range, ByVal formatKey As String)
    Const PROC As String = "ApplyFormat"
    On Error GoTo ErrHandler

    If rng Is Nothing Then
        Log PROC & ": rng é Nothing — abortando."
        Exit Sub
    End If

    Dim fmt As FormatDef
    fmt = GetFormatDef(formatKey)

    If fmt.key = "" Then
        Log PROC & ": chave desconhecida [" & formatKey & "]"
        MsgBox "Formato desconhecido: " & formatKey, vbExclamation, CFG_ADDIN_NAME
        Exit Sub
    End If

    Application.ScreenUpdating = False

    With rng
        .NumberFormat = fmt.NumberFmt
        If fmt.Alignment <> xlHAlignGeneral Then
            .HorizontalAlignment = fmt.Alignment
        End If
    End With

    Log PROC & ": aplicado [" & fmt.DisplayName & "] em " & rng.Address(External:=True)

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    HandleError PROC, Err
    Resume Cleanup
End Sub


' Sobrecarga conveniente: usa a seleção ativa.
Public Sub ApplyFormatToSelection(ByVal formatKey As String)
    Const PROC As String = "ApplyFormatToSelection"
    On Error GoTo ErrHandler

    Dim rng As Range
    Set rng = SafeSelection()
    If rng Is Nothing Then Exit Sub

    ApplyFormat rng, formatKey
    Exit Sub

ErrHandler:
    HandleError PROC, Err
End Sub


' Retorna o FormatDef para uma chave — útil para inspeção/debug.
Public Function GetFormatDef(ByVal key As String) As FormatDef
    Dim f As FormatDef

    ' -- REGISTRY DE FORMATOS -------------------------------------------------
    ' Para adicionar um novo formato: inclua um Case aqui e uma constante em
    ' modConfig. Nenhum outro arquivo precisa ser alterado.
    ' -------------------------------------------------------------------------
    Select Case key

        ' -- Família Fin (numérico / financeiro) ------------------------------
        ' Todos os formatos Fin respeitam CFG_FORCE_ALIGN e CFG_ZERO_DASH.

        Case FMT_INTEGER
            f.key = FMT_INTEGER
            f.DisplayName = "Financeiro 0 casas"
            f.NumberFmt = AccountingFmt(0, applyZeroDash:=CFG_ZERO_DASH)
            f.Category = "numeric"

        Case FMT_FIN_2D
            f.key = FMT_FIN_2D
            f.DisplayName = "Financeiro 2 casas"
            f.NumberFmt = AccountingFmt(2, applyZeroDash:=CFG_ZERO_DASH)
            f.Category = "numeric"

        Case FMT_FIN_4D
            f.key = FMT_FIN_4D
            f.DisplayName = "Financeiro 4 casas"
            f.NumberFmt = AccountingFmt(4, applyZeroDash:=CFG_ZERO_DASH)
            f.Category = "numeric"

        Case FMT_FIN_8D
            f.key = FMT_FIN_8D
            f.DisplayName = "Financeiro 8 casas"
            f.NumberFmt = AccountingFmt(8, applyZeroDash:=CFG_ZERO_DASH)
            f.Category = "numeric"

        ' -- Percentual -------------------------------------------------------
        Case FMT_PCT_4D
            f.key = FMT_PCT_4D
            f.DisplayName = "% 4 casas"
            f.NumberFmt = "0.0000%"
            f.Category = "percent"

        Case FMT_PCT_2D
            f.key = FMT_PCT_2D
            f.DisplayName = "% 2 casas"
            f.NumberFmt = "0.00%"
            f.Category = "percent"

        ' -- Spread em bps ----------------------------------------------------
        Case FMT_SPREAD_BPS
            f.key = FMT_SPREAD_BPS
            f.DisplayName = "Spread (bps)"
            f.NumberFmt = "#,##0.0"" bps"""
            f.Category = "numeric"

        ' -- Datas ------------------------------------------------------------
        Case FMT_DATE_ISO
            f.key = FMT_DATE_ISO
            f.DisplayName = "Data ISO"
            f.NumberFmt = "yyyy-mm-dd"
            f.Category = "date"

        Case FMT_DATE_BR
            f.key = FMT_DATE_BR
            f.DisplayName = "Data BR"
            f.NumberFmt = "[$-pt-BR]dd/mm/yyyy"
            f.Category = "date"

        Case FMT_DATE_BR_LONG
            f.key = FMT_DATE_BR_LONG
            f.DisplayName = "Data BR Longa"
            f.NumberFmt = "[$-pt-BR]dd/mmm/yyyy"
            f.Category = "date"

        ' -- Texto ------------------------------------------------------------
        Case FMT_TEXT
            f.key = FMT_TEXT
            f.DisplayName = "Texto"
            f.NumberFmt = "@"
            f.Category = "text"

        Case Else
            ' Retorna FormatDef vazio — o chamador detecta via f.Key = ""
            f.key = ""

    End Select

    GetFormatDef = f
End Function


' =============================================================================
' HELPERS PRIVADOS
' =============================================================================

' Gera string de formato contábil com N decimais — 3 seções: positivo;negativo;zero
'
' CFG_FORCE_ALIGN = True  ? inclui " * " (Excel repete espaços até preencher a célula)
' CFG_FORCE_ALIGN = False ? formato limpo, sem preenchimento de espaços
'
' applyZeroDash = True  ? seção zero usa _(-_)_-
'   _( abre espaço simétrico ao "(" do negativo
'   -  traço literal centralizado
'   _) fecha espaço simétrico ao ")" do negativo
'   _- espaço simétrico ao "-" final do positivo
' applyZeroDash = False ? seção zero repete o positivo (0,000... com casas)
Private Function AccountingFmt(ByVal decimals As Integer, _
                                Optional ByVal applyZeroDash As Boolean = False) As String
    Dim dec As String
    dec = String(decimals, "0")

    Dim pos As String, neg As String, zer As String

    If CFG_FORCE_ALIGN Then
        pos = " * _(#,##0." & dec & "_)_-"
        neg = " * (#,##0." & dec & ")_-"
        zer = IIf(applyZeroDash, " * _(-_)_-", pos)
    Else
        pos = "_(#,##0." & dec & "_)_-"
        neg = "(#,##0." & dec & ")_-"
        zer = IIf(applyZeroDash, "_(-_)_-", pos)
    End If

    ' Nota: decimais = 0 produz "_(#,##0_)_-" (sem ponto decimal)
    ' O String(0, "0") retorna "" — o ponto é omitido corretamente
    ' pois a concatenação "0." & "" resulta em "0." que não é o desejado.
    ' Portanto, trata-se o caso zero explicitamente:
    If decimals = 0 Then
        If CFG_FORCE_ALIGN Then
            pos = " * _(#,##0_)_-"
            neg = " * (#,##0)_-"
            zer = IIf(applyZeroDash, " * _(-_)_-", pos)
        Else
            pos = "_(#,##0_)_-"
            neg = "(#,##0)_-"
            zer = IIf(applyZeroDash, "_(-_)_-", pos)
        End If
    End If

    AccountingFmt = pos & ";" & neg & ";" & zer
End Function


