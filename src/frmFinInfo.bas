' =============================================================================
' frmFinInfo — UserForm: Guia de Referência dos Formatos Fin
'
' Propriedades do form (configurar no painel Properties do VBA):
'   Name            frmFinInfo
'   Caption         Formatos Financeiros — Guia de Referência
'   Width           540
'   Height          480
'   StartUpPosition 1 – CenterOwner
'
' Controles a inserir via toolbox:
'
'   Label (lblTitle)
'     Caption       Família Fin — Referência completa
'     FontBold      True
'     FontSize      10
'     Top           8
'     Left          10
'     Width         500
'
'   TextBox (txtContent)
'     MultiLine     True
'     ScrollBars    2 – fmScrollBarsVertical
'     FontName      Courier New
'     FontSize      9
'     Top           30
'     Left          8
'     Width         510
'     Height        370
'     Locked        True
'     TabStop       False
'
'   CommandButton (btnClose)
'     Caption       Fechar
'     Top           410
'     Left          420
'     Width         80
'     Height        24
'     Default       True
' =============================================================================

Option Explicit

Private Sub UserForm_Initialize()
    txtContent.Value = BuildFinInfoContent()
    ' Volta ao topo após preencher
    txtContent.SelStart = 0
    txtContent.SelLength = 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
