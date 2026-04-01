Option Explicit

' =============================================================================
' LÓGICA DO DIÁLOGO DE REFERÊNCIA DOS FORMATOS FIN
' =============================================================================

Public Sub ShowFinInfo()
    frmFinInfo.Show vbModal
End Sub


' Constrói o texto principal do diálogo.
' Usa valores hardcoded nos exemplos porque o caractere * do formato contábil
' não é suportado pela função Format() do VBA (é interpretado apenas pelo Excel).
Public Function BuildFinInfoContent() As String
    Const NL  As String = vbCrLf
    Const SEP As String = "  "
    Const W   As Integer = 64

    Dim s As String

    ' -- Exemplos -------------------------------------------------------------
    s = s & "Exemplos com o valor 1.234.567,8901234" & NL
    s = s & String(W, "=") & NL
    s = s & PadR("Formato", 10) & SEP & _
            PadR("Positivo", 20) & SEP & _
            PadR("Negativo", 20) & SEP & _
            "Zero" & NL
    s = s & String(W, "=") & NL

    Dim rows(3, 3) As String
    rows(0, 0) = "Fin 0D": rows(0, 1) = "1.234.568": rows(0, 2) = "(1.234.568)": rows(0, 3) = "0"
    rows(1, 0) = "Fin 2D": rows(1, 1) = "1.234.567,89": rows(1, 2) = "(1.234.567,89)": rows(1, 3) = "0,00"
    rows(2, 0) = "Fin 4D": rows(2, 1) = "1.234.567,8901": rows(2, 2) = "(1.234.567,8901)": rows(2, 3) = "0,0000"
    rows(3, 0) = "Fin 8D": rows(3, 1) = "1.234.567,8901": rows(3, 2) = "(1.234.567,8901)": rows(3, 3) = "0,00000000"

    Dim i As Integer
    For i = 0 To 3
        Dim zeroDisplay As String
        zeroDisplay = IIf(CFG_ZERO_DASH, "-", rows(i, 3))
        s = s & PadR(rows(i, 0), 10) & SEP & _
                PadR(rows(i, 1), 20) & SEP & _
                PadR(rows(i, 2), 20) & SEP & _
                zeroDisplay & NL
    Next i

    ' -- Configurações ativas -------------------------------------------------
    s = s & NL & String(W, "=") & NL
    s = s & "Configura" & Chr(231) & Chr(245) & "es ativas" & NL
    s = s & String(W, "=") & NL
    s = s & "  Alinhar " & Chr(224) & " direita (force) : " & BoolLabel(CFG_FORCE_ALIGN) & NL
    s = s & "  Zero como ""-""             : " & BoolLabel(CFG_ZERO_DASH) & NL

    ' -- Strings de formato ---------------------------------------------------
    s = s & NL & String(W, "=") & NL
    s = s & "Strings de formato (estado atual)" & NL
    s = s & String(W, "=") & NL

    Dim keys(3) As String
    keys(0) = FMT_INTEGER
    keys(1) = FMT_FIN_2D
    keys(2) = FMT_FIN_4D
    keys(3) = FMT_FIN_8D

    Dim sectionLabels(2) As String
    sectionLabels(0) = "pos"
    sectionLabels(1) = "neg"
    sectionLabels(2) = "zer"

    For i = 0 To 3
        Dim fmt As FormatDef
        fmt = GetFormatDef(keys(i))
        s = s & "  " & fmt.DisplayName & NL
        Dim sections() As String
        sections = Split(fmt.NumberFmt, ";")
        Dim j As Integer
        For j = 0 To UBound(sections)
            s = s & "    [" & sectionLabels(j) & "] " & sections(j) & NL
        Next j
        s = s & NL
    Next i

    ' -- Notas explicativas ---------------------------------------------------
    s = s & String(W, "=") & NL
    s = s & "Notas" & NL
    s = s & String(W, "=") & NL
    s = s & "  *  (asterisco)" & NL
    s = s & "     Excel repete o caractere seguinte ate preencher" & NL
    s = s & "     a largura da coluna. Controlado pelo checkbox" & NL
    s = s & "     ""Alinhar a direita (force)""." & NL & NL
    s = s & "  _( e _)" & NL
    s = s & "     Reserva um espaco equivalente aos parenteses" & NL
    s = s & "     do negativo, mantendo positivos e zeros" & NL
    s = s & "     alinhados com negativos na mesma coluna." & NL & NL
    s = s & "  _-" & NL
    s = s & "     Reserva um espaco equivalente ao sinal de menos," & NL
    s = s & "     alinhando positivos (sem sinal) com negativos." & NL & NL
    s = s & "  Secao zero com dash  _(-_)_-" & NL
    s = s & "     O traco fica no mesmo recuo dos digitos, entre" & NL
    s = s & "     _( e _), com _- final simetrico ao positivo." & NL

    BuildFinInfoContent = s
End Function


' -- Helpers privados ---------------------------------------------------------

Private Function PadR(ByVal s As String, ByVal width As Integer) As String
    If Len(s) >= width Then
        PadR = Left(s, width)
    Else
        PadR = s & Space(width - Len(s))
    End If
End Function

Private Function BoolLabel(ByVal b As Boolean) As String
    BoolLabel = IIf(b, "Sim", "Nao")
End Function


