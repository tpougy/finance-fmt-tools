#Requires -Version 5.1
<#
.SYNOPSIS
    Instala ou atualiza o add-in RBR Finance Tools no Excel.

.DESCRIPTION
    Baixa a versão mais recente do GitHub Releases e instala/atualiza
    o add-in no Excel. Detecta automaticamente se é uma instalação nova
    ou atualização de versão existente.

.PARAMETER Force
    Reinstala mesmo que o arquivo instalado pareça idêntico ao disponível.

.EXAMPLE
    # Execução direta (uma linha no PowerShell):
    Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/Install-RBRFinanceTools.ps1 | iex

.EXAMPLE
    # Execução local com flag de reinstalação forçada:
    .\Install-RBRFinanceTools.ps1 -Force
#>

[CmdletBinding()]
param(
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================================
# CONFIGURAÇÃO — ajuste se o repositório ou nome do arquivo mudarem
# =============================================================================
$CFG = @{
    AddinTitle    = 'Finance Fmt Tools'                   # Deve bater com o Title do .xlam (File > Info > Properties > Title)
    AddinFilename = 'FinanceFmtTools.xlam'              # Nome fixo do asset no GitHub Release
    GithubOwner   = 'tpougy'
    GithubRepo    = 'finance-fmt-tools'
}

$CFG.DownloadUrl  = "https://github.com/$($CFG.GithubOwner)/$($CFG.GithubRepo)/releases/latest/download/$($CFG.AddinFilename)"
$CFG.TempPath     = Join-Path $env:TEMP $CFG.AddinFilename
$CFG.OfficeAddins = Join-Path $env:APPDATA 'Microsoft\AddIns'
$CFG.DestPath     = Join-Path $CFG.OfficeAddins $CFG.AddinFilename


# =============================================================================
# HELPERS
# =============================================================================

function Write-Step {
    param([string]$Msg)
    Write-Host "  $Msg" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Msg)
    Write-Host "  [OK] $Msg" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Msg)
    Write-Host "  [!]  $Msg" -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Msg)
    Write-Host "  [X]  $Msg" -ForegroundColor Red
}

function Get-FileSizeKB {
    param([string]$Path)
    if (Test-Path $Path) { return [math]::Round((Get-Item $Path).Length / 1KB, 1) }
    return $null
}

function Assert-ExcelNotRunning {
    $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
    if ($excelProcs) {
        Write-Warn 'O Excel está aberto. Feche-o antes de instalar/atualizar o add-in.'
        $answer = Read-Host '  Deseja fechar o Excel automaticamente? [S/N]'
        if ($answer -match '^[Ss]') {
            $excelProcs | Stop-Process -Force
            Start-Sleep -Seconds 2
            Write-Ok 'Excel encerrado.'
        } else {
            throw 'Instalação cancelada. Feche o Excel e tente novamente.'
        }
    }
}

function Get-LatestReleaseTag {
    # Consulta a API do GitHub para obter a tag da versão mais recente
    $apiUrl = "https://api.github.com/repos/$($CFG.GithubOwner)/$($CFG.GithubRepo)/releases/latest"
    try {
        $release = Invoke-RestMethod -Uri $apiUrl -Headers @{ 'User-Agent' = 'RBR-Install' } -ErrorAction Stop
        return $release.tag_name
    } catch {
        return '(desconhecida)'
    }
}


# =============================================================================
# ETAPA 1 — DOWNLOAD COM BARRA DE PROGRESSO
# =============================================================================

function Get-FileFromWeb {
    param (
        [Parameter(Mandatory)]
        [string]$URL,

        [Parameter(Mandatory)]
        [string]$File
    )
    Begin {
        function Show-Progress {
            param (
                [Parameter(Mandatory)][Single]$TotalValue,
                [Parameter(Mandatory)][Single]$CurrentValue,
                [Parameter(Mandatory)][string]$ProgressText,
                [Parameter()][string]$ValueSuffix,
                [Parameter()][int]$BarSize = 40,
                [Parameter()][switch]$Complete
            )

            $percent = $CurrentValue / $TotalValue
            $percentComplete = $percent * 100
            if ($ValueSuffix) { $ValueSuffix = " $ValueSuffix" }

            if ($psISE) {
                Write-Progress "$ProgressText $CurrentValue$ValueSuffix of $TotalValue$ValueSuffix" -id 0 -percentComplete $percentComplete
            }
            else {
                $curBarSize = $BarSize * $percent
                $progbar = ""
                $progbar = $progbar.PadRight($curBarSize, [char]9608)
                $progbar = $progbar.PadRight($BarSize, [char]9617)

                if (!$Complete.IsPresent) {
                    Write-Host -NoNewLine "`r$ProgressText $progbar [ $($CurrentValue.ToString("#.###").PadLeft($TotalValue.ToString("#.###").Length))$ValueSuffix / $($TotalValue.ToString("#.###"))$ValueSuffix ] $($percentComplete.ToString("##0.00").PadLeft(6)) % complete"
                }
                else {
                    Write-Host -NoNewLine "`r$ProgressText $progbar [ $($TotalValue.ToString("#.###").PadLeft($TotalValue.ToString("#.###").Length))$ValueSuffix / $($TotalValue.ToString("#.###"))$ValueSuffix ] $($percentComplete.ToString("##0.00").PadLeft(6)) % complete"
                }
            }
        }
    }
    Process {
        try {
            $storeEAP = $ErrorActionPreference
            $ErrorActionPreference = 'Stop'

            $request  = [System.Net.HttpWebRequest]::Create($URL)
            $response = $request.GetResponse()

            if ($response.StatusCode -eq 401 -or $response.StatusCode -eq 403 -or $response.StatusCode -eq 404) {
                throw "Remote file either doesn't exist, is unauthorized, or is forbidden for '$URL'."
            }

            if ($File -match '^\.\\') {
                $File = Join-Path (Get-Location -PSProvider "FileSystem") ($File -Split '^\.')[1]
            }
            if ($File -and !(Split-Path $File)) {
                $File = Join-Path (Get-Location -PSProvider "FileSystem") $File
            }

            $fileDirectory = [System.IO.Path]::GetDirectoryName($File)
            if (!(Test-Path $fileDirectory)) {
                [System.IO.Directory]::CreateDirectory($fileDirectory) | Out-Null
            }

            [long]$fullSize   = $response.ContentLength
            $fullSizeMB       = $fullSize / 1024 / 1024
            [byte[]]$buffer   = New-Object byte[] 1048576
            [long]$total      = [long]$count = 0

            $reader = $response.GetResponseStream()
            $writer = New-Object System.IO.FileStream $File, "Create"

            # FIX: $File é string — usa GetFileName em vez de $File.Name (que seria FileInfo)
            $fileName     = [System.IO.Path]::GetFileName($File)
            $finalBarCount = 0

            do {
                $count = $reader.Read($buffer, 0, $buffer.Length)
                $writer.Write($buffer, 0, $count)
                $total += $count
                $totalMB = $total / 1024 / 1024

                if ($fullSize -gt 0) {
                    Show-Progress -TotalValue $fullSizeMB -CurrentValue $totalMB -ProgressText "Downloading $fileName" -ValueSuffix "MB"
                }

                if ($total -eq $fullSize -and $count -eq 0 -and $finalBarCount -eq 0) {
                    Show-Progress -TotalValue $fullSizeMB -CurrentValue $totalMB -ProgressText "Downloading $fileName" -ValueSuffix "MB" -Complete
                    $finalBarCount++
                }
            } while ($count -gt 0)
        }
        catch {
            $ExeptionMsg = $_.Exception.Message
            Write-Host "Download breaks with error : $ExeptionMsg"
        }
        finally {
            if ($reader) { $reader.Close() }
            if ($writer) { $writer.Flush(); $writer.Close() }
            $ErrorActionPreference = $storeEAP
            [GC]::Collect()
        }
    }
}

function Invoke-Download {
    Write-Step "Consultando versão mais recente em GitHub Releases..."
    $tag = Get-LatestReleaseTag
    Write-Step "Versão: $tag  |  URL: $($CFG.DownloadUrl)"

    # Remove arquivo temporário anterior se existir
    if (Test-Path $CFG.TempPath) { Remove-Item $CFG.TempPath -Force }

    Write-Step 'Baixando...'
    Write-Host ''   # linha em branco antes da barra inline

    Get-FileFromWeb -URL $CFG.DownloadUrl -File $CFG.TempPath

    Write-Host ''   # quebra de linha após a barra inline
    Write-Ok "Download concluído: $($CFG.TempPath)  ($(Get-FileSizeKB $CFG.TempPath) KB)"

    return $tag
}


# =============================================================================
# ETAPA 2 — INSTALL / UPDATE VIA EXCEL COM
# =============================================================================

function Invoke-AddinInstall {
    param([string]$ReleaseTag)

    # --- Garante pasta de add-ins do Office -----------------------------------
    if (-not (Test-Path $CFG.OfficeAddins)) {
        New-Item -ItemType Directory -Path $CFG.OfficeAddins -Force | Out-Null
        Write-Ok "Pasta de add-ins criada: $($CFG.OfficeAddins)"
    }

    # --- Compara arquivo baixado com o instalado ------------------------------
    $tempSize = (Get-Item $CFG.TempPath).Length
    $destExists = Test-Path $CFG.DestPath

    if ($destExists -and -not $Force) {
        $destSize = (Get-Item $CFG.DestPath).Length
        if ($tempSize -eq $destSize) {
            Write-Ok "O arquivo instalado já é idêntico ao release $ReleaseTag. Nada a fazer."
            Write-Warn "Use -Force para reinstalar mesmo assim."
            return
        }
        Write-Step "Tamanho diferente (instalado: $([math]::Round($destSize/1KB))KB  /  novo: $([math]::Round($tempSize/1KB))KB) — atualizando."
    } elseif (-not $destExists) {
        Write-Step 'Add-in não encontrado. Realizando instalação inicial.'
    } else {
        Write-Step '-Force ativo — reinstalando independentemente da versão.'
    }

    # --- Copia para a pasta de add-ins ----------------------------------------
    Write-Step "Copiando para $($CFG.DestPath)..."
    Copy-Item -Path $CFG.TempPath -Destination $CFG.DestPath -Force
    Write-Ok 'Arquivo copiado.'

    # --- Instala/atualiza via Excel COM ---------------------------------------
    Write-Step 'Abrindo Excel (COM) para registrar o add-in...'
    $excel = $null
    $wb    = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        # Excel.AddIns.Add() exige um workbook aberto (bug conhecido do Excel COM)
        $wb = $excel.Workbooks.Add()

        # Verifica se o add-in já está registrado (cenário de atualização)
        $existingAddin = $null
        for ($i = 1; $i -le $excel.AddIns.Count; $i++) {
            $ai = $excel.AddIns.Item($i)
            if ($ai.Title -eq $CFG.AddinTitle) {
                $existingAddin = $ai
                break
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ai) | Out-Null
        }

        if ($null -ne $existingAddin) {
            # UPDATE: desinstala a entrada antiga antes de reregistrar
            Write-Step "Add-in '$($CFG.AddinTitle)' já registrado — atualizando registro."
            $existingAddin.Installed = $false
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($existingAddin) | Out-Null
            $existingAddin = $null
        }

        # Adiciona e ativa o add-in a partir do caminho de destino definitivo
        $newAddin = $excel.AddIns.Add($CFG.DestPath, $false)
        $newAddin.Installed = $true
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($newAddin) | Out-Null

        Write-Ok "Add-in '$($CFG.AddinTitle)' registrado e ativado com sucesso."

    } finally {
        if ($wb)    { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
        if ($excel) { $excel.Quit();     [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


# =============================================================================
# ETAPA 3 — LIMPEZA
# =============================================================================

function Invoke-Cleanup {
    if (Test-Path $CFG.TempPath) {
        Remove-Item $CFG.TempPath -Force -ErrorAction SilentlyContinue
        Write-Step 'Arquivo temporário removido.'
    }
}


# =============================================================================
# PONTO DE ENTRADA
# =============================================================================

Write-Host ''
Write-Host '======================================================' -ForegroundColor White
Write-Host "  RBR Finance Tools — Instalador"                       -ForegroundColor White
Write-Host '======================================================' -ForegroundColor White
Write-Host ''

try {
    Assert-ExcelNotRunning

    $tag = Invoke-Download
    Write-Host ''

    Invoke-AddinInstall -ReleaseTag $tag
    Write-Host ''

    Invoke-Cleanup

    Write-Host ''
    Write-Host '======================================================' -ForegroundColor Green
    Write-Host "  Instalação concluída! Abra o Excel para usar a aba" -ForegroundColor Green
    Write-Host "  'Finance Fmt' no ribbon."                            -ForegroundColor Green
    Write-Host '======================================================' -ForegroundColor Green
    Write-Host ''

} catch {
    Write-Host ''
    Write-Fail "Erro durante a instalação: $_"
    Write-Host ''
    exit 1
}

