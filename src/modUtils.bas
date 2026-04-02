Option Explicit

Private Const CFG_XML_NS    As String = "urn:finance-fmt-tools"
Private Const CFG_XML_ROOT  As String = "<FmtConfig xmlns=""urn:finance-fmt-tools"">" & _
                                                 "<ForceAlign>true</ForceAlign>" & _
                                                 "<ZeroDash>false</ZeroDash>" & _
                                             "</FmtConfig>"


' =============================================================================
' UTILITÁRIOS GERAIS
' =============================================================================


' -- Logging ------------------------------------------------------------------

Public Sub Log(ByVal msg As String)
    If Not CFG_LOG_ENABLED Then Exit Sub

    Dim entry As String
    entry = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & msg

    ' Sempre grava no Immediate (Debug)
    Debug.Print entry

    ' Opcional: grava em aba oculta
    If CFG_LOG_TO_SHEET Then LogToSheet entry
End Sub

Private Sub LogToSheet(ByVal entry As String)
    On Error Resume Next   ' não deixa o log quebrar a funcionalidade principal

    Dim wb  As Workbook
    Dim ws  As Worksheet

    Set wb = ThisWorkbook

    Set ws = wb.Sheets(CFG_LOG_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = CFG_LOG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden
        ws.Cells(1, 1).value = "Timestamp"
        ws.Cells(1, 2).value = "Message"
    End If

    Dim nextRow As Long
    nextRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).value = entry

    On Error GoTo 0
End Sub


' -- Tratamento de erro centralizado ------------------------------------------

Public Sub HandleError(ByVal source As String, ByVal e As ErrObject)
    Dim msg As String
    msg = source & " | Err " & e.Number & ": " & e.Description

    Log "ERROR: " & msg

    #If DEBUG_MODE Then
        MsgBox msg, vbCritical, CFG_ADDIN_NAME & " – Erro"
    #End If
End Sub


' -- Seleção segura -----------------------------------------------------------

' Retorna a Selection como Range, ou Nothing com mensagem amigável.
Public Function SafeSelection() As Range
    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione um intervalo de células antes de aplicar a formatação.", _
               vbInformation, CFG_ADDIN_NAME
        Set SafeSelection = Nothing
        Exit Function
    End If

    Set SafeSelection = Selection
    Exit Function

ErrHandler:
    Set SafeSelection = Nothing
End Function


' -- Informações do Add-in ----------------------------------------------------

Public Sub ShowAbout()
    MsgBox CFG_ADDIN_NAME & " v" & CFG_ADDIN_VERSION & Chr(13) & Chr(13) & _
           "Formatação financeira padronizada para mercado de capitais." & Chr(13) & Chr(13) & _
           "Autor: Thomaz Pougy", _
           vbInformation, "Sobre"
End Sub


' =============================================================================
' PERSISTÊNCIA EM CustomXMLPart
'
' XML armazenado no .xlam:
'   <FmtConfig xmlns="urn:finance-fmt-tools">
'     <ForceAlign>true</ForceAlign>
'     <ZeroDash>false</ZeroDash>
'   </FmtConfig>
'
' O namespace próprio evita colisão com partes do Office ou outros add-ins.
' =============================================================================


' Lê todas as configurações do CustomXMLPart.
Public Sub LoadConfig()
    Const PROC As String = "LoadConfig"
    On Error GoTo ErrHandler

    Dim part As CustomXMLPart
    Set part = FindOrCreateXMLPart()

    CFG_FORCE_ALIGN = ReadBoolNode(part, "ForceAlign", defaultVal:=True)
    CFG_ZERO_DASH   = ReadBoolNode(part, "ZeroDash",   defaultVal:=False)

    Log PROC & ": ForceAlign=" & CFG_FORCE_ALIGN & _
               " | ZeroDash=" & CFG_ZERO_DASH & _
               " | fonte=CustomXMLPart"
    Exit Sub

ErrHandler:
    CFG_FORCE_ALIGN = True
    CFG_ZERO_DASH   = False
    Log PROC & ": ERRO na leitura — aplicando defaults | ForceAlign=True | ZeroDash=False | fonte=fallback"
    HandleError PROC, Err
End Sub


' Grava todas as configurações no CustomXMLPart.
Public Sub SaveConfig()
    Const PROC As String = "SaveConfig"
    On Error GoTo ErrHandler

    Dim part As CustomXMLPart
    Set part = FindOrCreateXMLPart()

    WriteBoolNode part, "ForceAlign", CFG_FORCE_ALIGN
    WriteBoolNode part, "ZeroDash",   CFG_ZERO_DASH

    ' Persiste em disco — sem este Save, as alterações no CustomXMLPart
    ' ficam apenas em memória e são perdidas quando o Excel fecha.
    Application.EnableEvents = False
    ThisWorkbook.Save
    Application.EnableEvents = True

    Log PROC & ": ForceAlign=" & CFG_FORCE_ALIGN & _
               " | ZeroDash=" & CFG_ZERO_DASH & _
               " | .xlam gravado em disco"
    Exit Sub

ErrHandler:
    Application.EnableEvents = True   ' reativa mesmo em erro
    Log PROC & ": ERRO ao salvar | ForceAlign=" & CFG_FORCE_ALIGN & _
               " | ZeroDash=" & CFG_ZERO_DASH
    HandleError PROC, Err
End Sub


' =============================================================================
' HELPERS PRIVADOS
' =============================================================================

' Lê um nó booleano; retorna defaultVal se o nó não existir (compatibilidade
' com arquivos salvos antes da adição do nó).
' Navega via DocumentElement/ChildNodes para evitar XPath com namespace,
' que não é suportado em todas as versões do Office.
Private Function ReadBoolNode(ByVal part As CustomXMLPart, _
                               ByVal nodeName As String, _
                               ByVal defaultVal As Boolean) As Boolean
    Dim node As CustomXMLNode
    Set node = FindChildNode(part, nodeName)

    If node Is Nothing Then
        ReadBoolNode = defaultVal
    Else
        ReadBoolNode = (LCase(Trim(node.Text)) = "true")
    End If
End Function


' Grava um nó booleano; se o nó não existir (arquivo legado), cria-o no root.
Private Sub WriteBoolNode(ByVal part As CustomXMLPart, _
                           ByVal nodeName As String, _
                           ByVal value As Boolean)
    Const PROC As String = "WriteBoolNode"
    On Error GoTo ErrHandler

    Dim node As CustomXMLNode
    Set node = FindChildNode(part, nodeName)

    ' Nó ausente ? arquivo legado; adiciona ao elemento raiz sem recriar tudo
    If node Is Nothing Then
        Dim root As CustomXMLNode
        Set root = part.DocumentElement
        If Not root Is Nothing Then
            root.AppendChildNode nodeName, CFG_XML_NS, msoCustomXMLNodeElement
            Set node = FindChildNode(part, nodeName)
        End If
    End If

    If Not node Is Nothing Then node.Text = LCase(CStr(value))
    Exit Sub

ErrHandler:
    HandleError PROC, Err
End Sub


' Localiza o CustomXMLPart pelo namespace; cria se não existir.
Private Function FindOrCreateXMLPart() As CustomXMLPart
    Const PROC As String = "FindOrCreateXMLPart"

    Dim i As Long
    For i = 1 To ThisWorkbook.CustomXMLParts.Count
        On Error Resume Next
        Dim ns As String
        ns = ThisWorkbook.CustomXMLParts(i).NamespaceURI
        On Error GoTo 0
        If ns = CFG_XML_NS Then
            Set FindOrCreateXMLPart = ThisWorkbook.CustomXMLParts(i)
            Log PROC & ": parte encontrada (índice " & i & ")"
            Exit Function
        End If
    Next i

    Set FindOrCreateXMLPart = ThisWorkbook.CustomXMLParts.Add(CFG_XML_ROOT)
    Log PROC & ": nova CustomXMLPart criada"
End Function


' Retorna o filho direto do elemento raiz cujo BaseName corresponde a nodeName.
' Substitui SelectSingleNode com namespace, evitando incompatibilidade entre
' versões do Office.
Private Function FindChildNode(ByVal part As CustomXMLPart, _
                                ByVal nodeName As String) As CustomXMLNode
    Set FindChildNode = Nothing

    Dim root As CustomXMLNode
    Set root = part.DocumentElement
    If root Is Nothing Then Exit Function

    Dim child As CustomXMLNode
    For Each child In root.ChildNodes
        If child.BaseName = nodeName Then
            Set FindChildNode = child
            Exit Function
        End If
    Next child
End Function

' -- Documentação online ------------------------------------------------------

Public Sub OpenDocsURL()
    Const PROC As String = "OpenDocsURL"
    On Error GoTo ErrHandler

    ThisWorkbook.FollowHyperlink CFG_DOCS_URL, NewWindow:=True
    Log PROC & ": abriu " & CFG_DOCS_URL
    Exit Sub

ErrHandler:
    HandleError PROC, Err
End Sub
