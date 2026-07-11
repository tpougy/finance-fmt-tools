#Requires -Version 5.1
<#
.SYNOPSIS
    Desinstalador do add-in COM "Finance Fmt Tools" (C# / Excel).

.DESCRIPTION
    Remove tudo o que o scripts\install.ps1 criou, SEM exigir administrador:
      1. PRE: confere se o Excel está FECHADO (com -Force, fecha por você).
      2. Remove as 3 árvores de registro em HKCU:
         - classe COM (CLSID + ProgId em HKCU\Software\Classes);
         - chave de descoberta (HKCU\...\Excel\Addins\FinanceFmtTools.Connect);
         - valor em ...\Resiliency\DoNotDisableAddinList (NUNCA a chave inteira —
           outros add-ins podem ter seus próprios valores ali).
      3. Remove os binários de %LocalAppData%\FinanceFmtTools\.
      4. RELATÓRIO do que foi removido.

    IDEMPOTENTE: rodar de novo (mesmo já desinstalado) NÃO falha — exit 0
    incondicional ao final, exceto quando o Excel está aberto e -Force não
    foi informado.

.PARAMETER Force
    Fecha o Excel automaticamente (com aviso) se estiver aberto.

.EXAMPLE
    Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/uninstall.ps1 | iex

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1 -Force

.NOTES
    Compatível com Windows PowerShell 5.1+. Sem admin, sem regasm.
#>

[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# Identidade fixa (mesmos valores de scripts/install.ps1 — declarada de forma
# independente aqui; este script não depende do arquivo install.ps1)
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
        Write-Ok 'Excel está fechado.'
        return
    }

    if ($Force) {
        Write-Warn2 'Excel está aberto. -Force informado: tentando fechar com segurança...'
        try {
            $excelProcs | ForEach-Object { $_.CloseMainWindow() | Out-Null }
            Start-Sleep -Seconds 3
            $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
            if ($excelProcs) {
                Write-Warn2 'Excel não fechou sozinho; encerrando o processo (-Force)...'
                $excelProcs | Stop-Process -Force -ErrorAction Stop
                Start-Sleep -Seconds 2
            }
            Write-Ok 'Excel fechado.'
        } catch {
            Write-Err2 ("Não consegui fechar o Excel automaticamente: {0}" -f $_.Exception.Message)
            exit 1
        }
    } else {
        Write-Err2 'O Excel está ABERTO. Feche-o completamente antes de desinstalar, ou rode novamente com -Force.'
        exit 1
    }
}

# Remove uma árvore de registro se existir; nunca falha se já removida (idempotente).
function Remove-KeyIfExists {
    param([string]$Path)
    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        Write-Ok ("Removida árvore de registro: {0}" -f $Path)
    } else {
        Write-Info ("Já ausente (nada a remover): {0}" -f $Path)
    }
}

Write-Host ''
Write-Host '############################################################' -ForegroundColor White
Write-Host '#  Desinstalador - Finance Fmt Tools (C#)  v1.0.0         #' -ForegroundColor White
Write-Host '############################################################' -ForegroundColor White

# ===========================================================================
# PASSO 1 - Pré-remoção: Excel fechado?
# ===========================================================================
Write-Step 'Pré-remoção'
Assert-ExcelNotRunning

# ===========================================================================
# PASSO 2 - Remover as 3 árvores de registro (HKCU)
# ===========================================================================
Write-Step 'Removendo chaves de registro (HKCU)'

# (a) Classe COM: CLSID (inclui ProgId + InprocServer32 como filhos) + ProgId->CLSID
Remove-KeyIfExists -Path "HKCU:\Software\Classes\CLSID\$Guid"
Remove-KeyIfExists -Path "HKCU:\Software\Classes\$ProgId"

# (b) Descoberta pelo Excel (NÃO versionado)
Remove-KeyIfExists -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgId"

# (c) Resiliência: remover apenas o VALOR do ProgId — a chave DoNotDisableAddinList
# pode conter valores de outros add-ins e NUNCA deve ser removida por inteiro.
$kResil = "HKCU:\Software\Microsoft\Office\$OfficeVerKey\Excel\Resiliency\DoNotDisableAddinList"
if (Test-Path $kResil) {
    $prop = Get-ItemProperty -Path $kResil -Name $ProgId -ErrorAction SilentlyContinue
    if ($null -ne $prop -and $null -ne $prop.$ProgId) {
        Remove-ItemProperty -Path $kResil -Name $ProgId -Force -ErrorAction Stop
        Write-Ok ("Removido valor de resiliência: {0}\{1}" -f $kResil, $ProgId)
    } else {
        Write-Info ("Valor de resiliência já ausente: {0}\{1}" -f $kResil, $ProgId)
    }
} else {
    Write-Info ("Chave de resiliência já ausente: {0}" -f $kResil)
}

# ===========================================================================
# PASSO 3 - Remover arquivos instalados
# ===========================================================================
Write-Step 'Removendo arquivos instalados'

if (Test-Path $InstallDir) {
    foreach ($f in $AllFiles) {
        $p = Join-Path $InstallDir $f
        if (Test-Path $p) {
            Remove-Item -Path $p -Force -ErrorAction Stop
            Write-Ok ("Removido: {0}" -f $f)
        } else {
            Write-Info ("Já ausente: {0}" -f $f)
        }
    }
    # Remove a pasta de instalação somente se ficou vazia (nunca um delete em bloco —
    # arquivos não listados podem ter sido colocados ali manualmente).
    $resto = Get-ChildItem -Path $InstallDir -Force -ErrorAction SilentlyContinue
    if (-not $resto) {
        Remove-Item -Path $InstallDir -Force -ErrorAction SilentlyContinue
        Write-Info ("Pasta de instalação vazia removida: {0}" -f $InstallDir)
    }
} else {
    Write-Info ("Pasta de instalação já ausente: {0}" -f $InstallDir)
}

# ===========================================================================
# PASSO 4 - Relatório final
# ===========================================================================
Write-Step 'Relatório final'
Write-Ok 'Desinstalação concluída.'
Write-Host ''
Write-Host 'Removido (ou já ausente):' -ForegroundColor White
Write-Host ("  - Classe COM (CLSID {0} + ProgId {1}) em HKCU\Software\Classes" -f $Guid, $ProgId)
Write-Host ("  - Chave de add-in: HKCU\Software\Microsoft\Office\Excel\Addins\{0}" -f $ProgId)
Write-Host ("  - Valor de resiliência: ...\DoNotDisableAddinList\{0}" -f $ProgId)
Write-Host ("  - Binários do add-in em {0}" -f $InstallDir)
Write-Host ''

exit 0
