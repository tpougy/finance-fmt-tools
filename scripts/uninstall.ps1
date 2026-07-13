#Requires -Version 5.1
<#
.SYNOPSIS
    Desinstalador do add-in COM "Finance Fmt Tools" (C# / Excel).

.DESCRIPTION
    Remove tudo o que o scripts\install.ps1 criou, SEM exigir administrador:
      1. PRE: confere se o Excel esta FECHADO (com -Force, fecha por voce).
      2. Remove as 3 arvores de registro em HKCU:
         - classe COM (CLSID + ProgId em HKCU\Software\Classes);
         - chave de descoberta (HKCU\...\Excel\Addins\FinanceFmtTools.Connect);
         - valor em ...\Resiliency\DoNotDisableAddinList (NUNCA a chave inteira -
           outros add-ins podem ter seus proprios valores ali).
      3. Remove os binarios de %LocalAppData%\FinanceFmtTools\.
      4. RELATORIO do que foi removido.

    IDEMPOTENTE: rodar de novo (mesmo ja desinstalado) NAO falha - exit 0
    incondicional ao final, exceto quando o Excel esta aberto e -Force nao
    foi informado.

.PARAMETER Force
    Fecha o Excel automaticamente (com aviso) se estiver aberto.

.EXAMPLE
    Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/uninstall.ps1 | iex

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1 -Force

.NOTES
    Compativel com Windows PowerShell 5.1+. Sem admin, sem regasm.
#>

[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# Identidade fixa (mesmos valores de scripts/install.ps1 - declarada de forma
# independente aqui; este script nao depende do arquivo install.ps1)
# =============================================================================
$Guid         = '{881EFDF3-424C-4240-BCA0-714DAC2B9CD7}'
$ProgId       = 'FinanceFmtTools.Connect'
$OfficeVerKey = '16.0'

$InstallDir = Join-Path $env:LOCALAPPDATA 'FinanceFmtTools'
$AllFiles   = @('FinanceFmtTools.ComAddin.dll', 'FinanceFmtTools.Engine.dll', 'Microsoft.Office.Interop.Excel.dll', 'office.dll')

function Write-Step  { param([string]$m) Write-Host ''; Write-Host ('=== {0} ===' -f $m) -ForegroundColor White }
function Write-Ok    { param([string]$m) Write-Host ('[OK]      ' + $m) -ForegroundColor Green }
function Write-Info  { param([string]$m) Write-Host ('[INFO]    ' + $m) -ForegroundColor Cyan }
function Write-Warn2 { param([string]$m) Write-Host ('[AVISO]   ' + $m) -ForegroundColor Yellow }
function Write-Err2  { param([string]$m) Write-Host ('[ERRO]    ' + $m) -ForegroundColor Red }

# Fecha o Excel (ou orienta a fechar) antes de tocar em qualquer arquivo/registro.
function Assert-ExcelNotRunning {
    $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
    if (-not $excelProcs) {
        Write-Ok 'Excel esta fechado.'
        return
    }

    if ($Force) {
        # NUNCA forca o encerramento do processo (Stop-Process): CloseMainWindow() pode
        # abrir um dialogo nativo "Salvar alteracoes?" para QUALQUER pasta de trabalho
        # aberta, nao so a deste desinstalador. Matar o processo descartaria esse dialogo
        # e o trabalho nao salvo do usuario. Em vez disso, aguarda ate 30s pelo fechamento
        # espontaneo e falha com uma mensagem acionavel se o Excel continuar aberto.
        Write-Warn2 'Excel esta aberto. -Force informado: solicitando fechamento (ate 30s)...'
        try {
            $excelProcs | ForEach-Object { $_.CloseMainWindow() | Out-Null }
            for ($i = 0; $i -lt 30; $i++) {
                Start-Sleep -Seconds 1
                $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
                if (-not $excelProcs) { break }
            }
            if ($excelProcs) {
                Write-Err2 'Excel ainda esta aberto (pode haver um dialogo "Salvar alteracoes?" pendente). Salve seu trabalho, feche o Excel manualmente e rode o desinstalador novamente.'
                exit 1
            }
            Write-Ok 'Excel fechado.'
        } catch {
            Write-Err2 ("Nao consegui fechar o Excel automaticamente: {0}" -f $_.Exception.Message)
            exit 1
        }
    } else {
        Write-Err2 'O Excel esta ABERTO. Feche-o completamente antes de desinstalar, ou rode novamente com -Force.'
        exit 1
    }
}

# Remove uma arvore de registro se existir; nunca falha se ja removida (idempotente).
function Remove-KeyIfExists {
    param([string]$Path)
    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        Write-Ok ("Removida arvore de registro: {0}" -f $Path)
    } else {
        Write-Info ("Ja ausente (nada a remover): {0}" -f $Path)
    }
}

Write-Host ''
Write-Host '############################################################' -ForegroundColor White
Write-Host '#  Desinstalador - Finance Fmt Tools (C#)  v1.0.0         #' -ForegroundColor White
Write-Host '############################################################' -ForegroundColor White

# ===========================================================================
# PASSO 1 - Pre-remocao: Excel fechado?
# ===========================================================================
Write-Step 'Pre-remocao'
Assert-ExcelNotRunning

try {
    # ===========================================================================
    # PASSO 2 - Remover as 3 arvores de registro (HKCU)
    # ===========================================================================
    Write-Step 'Removendo chaves de registro (HKCU)'

    # (a) Classe COM: CLSID (inclui ProgId + InprocServer32 como filhos) + ProgId->CLSID
    Remove-KeyIfExists -Path "HKCU:\Software\Classes\CLSID\$Guid"
    Remove-KeyIfExists -Path "HKCU:\Software\Classes\$ProgId"

    # (b) Descoberta pelo Excel (NAO versionado)
    Remove-KeyIfExists -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgId"

    # (c) Resiliencia: remover apenas o VALOR do ProgId - a chave DoNotDisableAddinList
    # pode conter valores de outros add-ins e NUNCA deve ser removida por inteiro.
    $kResil = "HKCU:\Software\Microsoft\Office\$OfficeVerKey\Excel\Resiliency\DoNotDisableAddinList"
    if (Test-Path $kResil) {
        $prop = Get-ItemProperty -Path $kResil -Name $ProgId -ErrorAction SilentlyContinue
        if ($null -ne $prop -and $null -ne $prop.$ProgId) {
            Remove-ItemProperty -Path $kResil -Name $ProgId -Force -ErrorAction Stop
            Write-Ok ("Removido valor de resiliencia: {0}\{1}" -f $kResil, $ProgId)
        } else {
            Write-Info ("Valor de resiliencia ja ausente: {0}\{1}" -f $kResil, $ProgId)
        }
    } else {
        Write-Info ("Chave de resiliencia ja ausente: {0}" -f $kResil)
    }

    # ===========================================================================
    # PASSO 3 - Remover arquivos instalados
    # ===========================================================================
    Write-Step 'Removendo arquivos instalados'

    # Reconfere que o Excel continua fechado imediatamente antes de remover
    # arquivos - evita um file-lock silencioso caso o usuario tenha reaberto o
    # Excel entre a checagem inicial (Passo 1) e este ponto (TOCTOU).
    Assert-ExcelNotRunning

    if (Test-Path $InstallDir) {
        foreach ($f in $AllFiles) {
            $p = Join-Path $InstallDir $f
            if (Test-Path $p) {
                Remove-Item -Path $p -Force -ErrorAction Stop
                Write-Ok ("Removido: {0}" -f $f)
            } else {
                Write-Info ("Ja ausente: {0}" -f $f)
            }
        }
        # Remove a pasta de instalacao somente se ficou vazia (nunca um delete em bloco -
        # arquivos nao listados podem ter sido colocados ali manualmente).
        $resto = Get-ChildItem -Path $InstallDir -Force -ErrorAction SilentlyContinue
        if (-not $resto) {
            Remove-Item -Path $InstallDir -Force -ErrorAction SilentlyContinue
            Write-Info ("Pasta de instalacao vazia removida: {0}" -f $InstallDir)
        }
    } else {
        Write-Info ("Pasta de instalacao ja ausente: {0}" -f $InstallDir)
    }
} catch {
    Write-Err2 ("Falha durante a desinstalacao (registro ou arquivos): {0}" -f $_.Exception.Message)
    exit 1
}

# ===========================================================================
# PASSO 4 - Relatorio final
# ===========================================================================
Write-Step 'Relatorio final'
Write-Ok 'Desinstalacao concluida.'
Write-Host ''
Write-Host 'Removido (ou ja ausente):' -ForegroundColor White
Write-Host ("  - Classe COM (CLSID {0} + ProgId {1}) em HKCU\Software\Classes" -f $Guid, $ProgId)
Write-Host ("  - Chave de add-in: HKCU\Software\Microsoft\Office\Excel\Addins\{0}" -f $ProgId)
Write-Host ("  - Valor de resiliencia: ...\DoNotDisableAddinList\{0}" -f $ProgId)
Write-Host ("  - Binarios do add-in em {0}" -f $InstallDir)
Write-Host ''

exit 0
