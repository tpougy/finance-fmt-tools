Option Explicit

' -- Identidade do Add-in -----------------------------------------------------
Public Const CFG_ADDIN_NAME      As String = "Finance Fmt Tools"
Public Const CFG_ADDIN_VERSION   As String = "1.3.0"
Public Const CFG_RIBBON_TAB_ID   As String = "tabFinanceFmt"

' -- Logging ------------------------------------------------------------------
Public Const CFG_LOG_ENABLED     As Boolean = True
Public Const CFG_LOG_TO_SHEET    As Boolean = False   ' True ? grava em aba oculta
Public Const CFG_LOG_SHEET_NAME  As String = "_FTLog"

' -- Estado persistido (carregado por LoadConfig em OnRibbonLoad) -------------
Public CFG_FORCE_ALIGN           As Boolean  ' True  ? * preenche coluna com espaços
Public CFG_ZERO_DASH             As Boolean  ' True  ? zero exibido como "-"

' -- Chaves de formato (evita magic strings em todo o projeto) -----------------
Public Const FMT_FIN_8D          As String = "FIN_8D"
Public Const FMT_FIN_4D          As String = "FIN_4D"
Public Const FMT_FIN_2D          As String = "FIN_2D"
Public Const FMT_PCT_4D          As String = "PCT_4D"
Public Const FMT_PCT_2D          As String = "PCT_2D"
Public Const FMT_SPREAD_BPS      As String = "SPREAD_BPS"
Public Const FMT_DATE_ISO        As String = "DATE_ISO"
Public Const FMT_DATE_BR         As String = "DATE_BR"
Public Const FMT_DATE_BR_LONG    As String = "DATE_BR_LONG"
Public Const FMT_TEXT            As String = "TEXT"
Public Const FMT_INTEGER         As String = "INTEGER"

