Option Explicit

' =============================================================================
' RIBBON CALLBACKS
' Regra: cada callback tem exatamente 1 linha de lógica.
' Toda lógica real está em modFormatEngine, modUtils e modFinInfo.
' =============================================================================

' Referência ao ribbon — necessária para getPressed funcionar
Private mRibbon As IRibbonUI


' -- Inicialização ------------------------------------------------------------

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    LoadConfig      ' lê CFG_FORCE_ALIGN e CFG_ZERO_DASH do CustomXMLPart
    Log "OnRibbonLoad: ForceAlign=" & CFG_FORCE_ALIGN & " | ZeroDash=" & CFG_ZERO_DASH
End Sub


' -- Família Fin --------------------------------------------------------------

Public Sub RibbonInteger(control As IRibbonControl)
    ApplyFormatToSelection FMT_INTEGER
End Sub

Public Sub RibbonFin2D(control As IRibbonControl)
    ApplyFormatToSelection FMT_FIN_2D
End Sub

Public Sub RibbonFin4D(control As IRibbonControl)
    ApplyFormatToSelection FMT_FIN_4D
End Sub

Public Sub RibbonFin8D(control As IRibbonControl)
    ApplyFormatToSelection FMT_FIN_8D
End Sub


' -- Percentual ---------------------------------------------------------------

Public Sub RibbonPct4D(control As IRibbonControl)
    ApplyFormatToSelection FMT_PCT_4D
End Sub

Public Sub RibbonPct2D(control As IRibbonControl)
    ApplyFormatToSelection FMT_PCT_2D
End Sub

Public Sub RibbonSpreadBps(control As IRibbonControl)
    ApplyFormatToSelection FMT_SPREAD_BPS
End Sub


' -- Datas --------------------------------------------------------------------

Public Sub RibbonDateISO(control As IRibbonControl)
    ApplyFormatToSelection FMT_DATE_ISO
End Sub

Public Sub RibbonDateBR(control As IRibbonControl)
    ApplyFormatToSelection FMT_DATE_BR
End Sub

Public Sub RibbonDateBRLong(control As IRibbonControl)
    ApplyFormatToSelection FMT_DATE_BR_LONG
End Sub


' -- Texto --------------------------------------------------------------------

Public Sub RibbonText(control As IRibbonControl)
    ApplyFormatToSelection FMT_TEXT
End Sub


' -- Checkboxes de configuração -----------------------------------------------

Public Sub RibbonChkForceAlign(control As IRibbonControl, pressed As Boolean)
    CFG_FORCE_ALIGN = pressed
    SaveConfig
    Log "ForceAlign alterado para " & pressed
End Sub

Public Sub RibbonGetForceAlign(control As IRibbonControl, ByRef returnValue As Variant)
    returnValue = CFG_FORCE_ALIGN
End Sub

Public Sub RibbonChkZeroDash(control As IRibbonControl, pressed As Boolean)
    CFG_ZERO_DASH = pressed
    SaveConfig
    Log "ZeroDash alterado para " & pressed
End Sub

Public Sub RibbonGetZeroDash(control As IRibbonControl, ByRef returnValue As Variant)
    returnValue = CFG_ZERO_DASH
End Sub


' -- Info ---------------------------------------------------------------------

Public Sub RibbonFinInfo(control As IRibbonControl)
    ShowFinInfo
End Sub

Public Sub RibbonAbout(control As IRibbonControl)
    ShowAbout
End Sub

