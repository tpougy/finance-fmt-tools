#Requires -Version 5.1
<#
.SYNOPSIS
    Instala o add-in COM "Finance Fmt Tools" (C# / Excel) no Excel 64-bit. Uso do USUARIO FINAL.
    Tambem detecta e remove automaticamente uma instalacao legada da versao VBA (.xlam),
    se presente, antes de instalar a versao C#.

.DESCRIPTION
    Registra o add-in COM FinanceFmtTools.ComAddin inteiramente em HKCU, sem exigir
    privilegios de administrador e sem usar regasm.exe. Excel 64-bit e o baseline
    validado (FUT-01 adia suporte/teste de 32-bit, nao necessariamente compatibilidade
    tecnica).

    FLUXO PRINCIPAL (nenhum parametro — o one-liner documentado):
      1. Detecta uma instalacao legada da versao VBA (FinanceFmtTools.xlam em
         %APPDATA%\Microsoft\AddIns) e, se encontrada, desregistra-a do Excel via
         automacao COM e remove o arquivo, antes de prosseguir com os passos
         abaixo.
      2. Baixa o pacote .zip mais recente do GitHub Releases (release-agnostic —
         a URL "latest/download/" nunca muda entre versoes).
      3. Extrai o .zip para uma pasta temporaria sob %TEMP% (nunca extrai direto
         para a pasta de instalacao final — mitigacao de zip-slip).
      4. Copia os 4 arquivos necessarios para %LocalAppData%\FinanceFmtTools\.
      5. Registra as 3 arvores de registro HKCU (classe COM, descoberta pelo Excel
         com LoadBehavior=3, e a chave de Resiliency DoNotDisableAddinList).
      6. Valida o pos-instalacao e limpa a pasta temporaria.

    ESCOTILHA DE TESTE LOCAL (-Package/-Source):
      Ate o Phase 5 publicar um release C# real via CI, use -Package <zip> ou
      -Source <pasta bin> para registrar a partir de um build local
      (dotnet build), sem contato com o GitHub.

    IDEMPOTENTE: rodar de novo sobrescreve os valores; nao corrompe nada.

.PARAMETER Package
    Caminho de um .zip local com os binarios (escotilha de teste local — 04-RESEARCH.md
    Pattern 2). Extrai o .zip e localiza a pasta com FinanceFmtTools.ComAddin.dll dentro.

.PARAMETER Source
    Pasta local (ja extraida ou pasta bin\ de um build) que contem os binarios.
    Escotilha de teste local, alternativa ao -Package.

.PARAMETER Force
    Fecha o Excel automaticamente (com aviso) se ele estiver aberto, em vez de
    apenas pedir que o usuario feche. Default: NAO forca (so orienta).

.EXAMPLE
    # Execucao direta (uma linha no PowerShell) — fluxo documentado, INST-01:
    Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/install.ps1 | iex

.EXAMPLE
    # Teste local a partir de um build (escotilha de teste local, sem GitHub):
    powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Source .\src\FinanceFmtTools.ComAddin\bin\Debug\net48

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Package .\FinanceFmtTools.zip -Force

.NOTES
    Compativel com Windows PowerShell 5.1+. NAO exige admin. NAO usa regasm.
    NAO escreve em HKLM (apenas leitura informativa de bitness do Office).
#>

[CmdletBinding()]
param(
    [string]$Package,
    [string]$Source,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# Identidade fixa (deve bater com src/FinanceFmtTools.ComAddin/Connect.cs — NAO
# inventar novos valores; ler do header doc-comment de Connect.cs e do .csproj)
# =============================================================================
$Guid         = '{881EFDF3-424C-4240-BCA0-714DAC2B9CD7}'
$ProgId       = 'FinanceFmtTools.Connect'
$ClassName    = 'FinanceFmtTools.ComAddin.Connect'
# $AssemblyStr NAO e' um literal fixo aqui: e' lido do próprio DLL copiado (via
# System.Reflection.AssemblyName), para nunca ficar dessincronizado de um bump de versão.
$RuntimeVer   = 'v4.0.30319'
$Shim         = 'C:\Windows\System32\mscoree.dll'
$ThreadingMdl = 'Both'
$FriendlyName = 'Finance Fmt Tools'
$Description  = 'Formatação financeira padronizada para mercado de capitais.'
$OfficeVerKey = '16.0'

# =============================================================================
# Legado VBA (.xlam) — deteccao/remocao automatica antes de instalar a versao C#
# =============================================================================
# $VbaAddinTitle deve bater exatamente com o document property Title do .xlam
# legado (usado para casar o add-in na colecao Excel.AddIns) — nao inventar
# outro valor.
$VbaAddinTitle = 'Finance Fmt Tools'
$VbaAddinDir   = Join-Path $env:APPDATA 'Microsoft\AddIns'
$VbaXlamPath   = Join-Path $VbaAddinDir 'FinanceFmtTools.xlam'

# =============================================================================
# GitHub Releases (INST-01) — URL "latest" e independente de versao de proposito
# =============================================================================
$GithubOwner = 'tpougy'
$GithubRepo  = 'finance-fmt-tools'
# Convencao introduzida por este plano: todo release do Phase 5 (CI) deve publicar
# seu zip sob este nome literal fixo, para que a URL "latest/download/" nunca
# precise mudar entre versoes — espelha o padrao de nome-fixo do instalador legado.
$AssetName   = 'FinanceFmtTools.zip'
$DownloadUrl = "https://github.com/$GithubOwner/$GithubRepo/releases/latest/download/$AssetName"

# =============================================================================
# Layout de arquivos instalados
# =============================================================================
$InstallDir = Join-Path $env:LOCALAPPDATA 'FinanceFmtTools'
$DllName    = 'FinanceFmtTools.ComAddin.dll'
$OtherFiles = @('FinanceFmtTools.Engine.dll', 'Microsoft.Office.Interop.Excel.dll', 'office.dll')
$AllFiles   = @($DllName) + $OtherFiles

# Forca TLS 1.2 — PS 5.1 no Windows usa TLS 1.0 por padrao, rejeitado pelo GitHub
# desde 2018. Deve ficar antes de qualquer chamada de rede.
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Pasta temporaria de extracao (limpa ao final, se criada).
$script:TempExtractDir = $null

# So vira $true quando uma instalacao VBA legada foi efetivamente detectada E
# removida do disco (consumido no relatorio final do PASSO 4).
$script:VbaRemoved = $false

# ---------------------------------------------------------------------------
# Saida formatada
# ---------------------------------------------------------------------------
function Write-Step  { param([string]$m) Write-Host ''; Write-Host ('=== {0} ===' -f $m) -ForegroundColor White }
function Write-Ok    { param([string]$m) Write-Host ('[OK]      ' + $m) -ForegroundColor Green }
function Write-Info  { param([string]$m) Write-Host ('[INFO]    ' + $m) -ForegroundColor Cyan }
function Write-Warn2 { param([string]$m) Write-Host ('[AVISO]   ' + $m) -ForegroundColor Yellow }
function Write-Err2  { param([string]$m) Write-Host ('[ERRO]    ' + $m) -ForegroundColor Red }

# Fecha o Excel (ou orienta a fechar) antes de tocar em qualquer arquivo/registro —
# evita o file-lock classico do DLL instalado durante uma reinstalacao.
function Assert-ExcelNotRunning {
    $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
    if (-not $excelProcs) {
        Write-Ok 'Excel está fechado.'
        return
    }

    if ($Force) {
        # NUNCA forca o encerramento do processo (Stop-Process): CloseMainWindow() pode
        # abrir um dialogo nativo "Salvar alteracoes?" para QUALQUER pasta de trabalho
        # aberta, nao so a deste instalador. Matar o processo descartaria esse diálogo e
        # o trabalho não salvo do usuário. Em vez disso, aguarda até 30s pelo fechamento
        # espontâneo e falha com uma mensagem acionável se o Excel continuar aberto.
        Write-Warn2 'Excel está aberto. -Force informado: solicitando fechamento (até 30s)...'
        try {
            $excelProcs | ForEach-Object { $_.CloseMainWindow() | Out-Null }
            for ($i = 0; $i -lt 30; $i++) {
                Start-Sleep -Seconds 1
                $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
                if (-not $excelProcs) { break }
            }
            if ($excelProcs) {
                Write-Err2 'Excel ainda está aberto (pode haver um diálogo "Salvar alterações?" pendente). Salve seu trabalho, feche o Excel manualmente e rode o instalador novamente.'
                exit 1
            }
            Write-Ok 'Excel fechado.'
        } catch {
            Write-Err2 ("Não consegui fechar o Excel automaticamente: {0}" -f $_.Exception.Message)
            exit 1
        }
    } else {
        Write-Err2 'O Excel está ABERTO. Feche-o completamente antes de instalar, ou rode novamente com -Force.'
        Write-Info 'Ex.: powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Force'
        exit 1
    }
}

# Detecta uma instalacao legada da versao VBA (.xlam), desregistra-a do Excel via
# automacao COM e remove o arquivo do disco. Nunca bloqueia a instalacao C#: qualquer
# falha na automacao COM apenas gera um aviso e a funcao segue em frente.
function Remove-LegacyVbaAddin {
    if (-not (Test-Path -LiteralPath $VbaXlamPath)) {
        # Sem instalacao legada — retorna cedo, sem abrir o Excel.
        return
    }
    Write-Info ("Instalação legada da versão VBA detectada: {0}" -f $VbaXlamPath)

    $excel      = $null
    $wb         = $null
    $foundAddin = $null

    try {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            # Necessario para acessar a colecao AddIns (mesmo padrao do instalador
            # VBA legado arquivado — ver contexto do plano).
            $wb = $excel.Workbooks.Add()

            for ($i = 1; $i -le $excel.AddIns.Count; $i++) {
                $ai = $excel.AddIns.Item($i)
                if ($ai.Title -eq $VbaAddinTitle) {
                    $foundAddin = $ai
                    break
                } else {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ai) | Out-Null
                }
            }

            if ($foundAddin) {
                $foundAddin.Installed = $false
                Write-Ok ("Add-in VBA legado '{0}' desregistrado do Excel." -f $VbaAddinTitle)
            } else {
                Write-Info ("Nenhum add-in registrado com o título '{0}' foi encontrado no Excel — o arquivo será removido mesmo assim." -f $VbaAddinTitle)
            }
        } catch {
            Write-Warn2 ("Não foi possível desregistrar o add-in VBA legado via Excel COM: {0}. O arquivo será removido mesmo assim." -f $_.Exception.Message)
        }
    } finally {
        if ($wb) {
            try { $wb.Close($false) } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        }
        if ($foundAddin) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($foundAddin) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    Remove-Item -LiteralPath $VbaXlamPath -Force -ErrorAction SilentlyContinue
    Write-Ok ("Instalação legada VBA (.xlam) removida: {0}" -f $VbaXlamPath)
    $script:VbaRemoved = $true
}

# Le o cabecalho PE (MS-DOS/COFF) e retorna 'x64' / 'x86' / hex / 'desconhecido'.
function Test-PeMachine {
    param([string]$Path)
    $fs = $null
    $br = $null
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $br = New-Object System.IO.BinaryReader($fs)
        $fs.Seek(0x3C, 'Begin') | Out-Null
        $peOff = $br.ReadInt32()
        $fs.Seek($peOff + 4, 'Begin') | Out-Null
        $machine = $br.ReadUInt16()
        switch ($machine) {
            0x8664  { 'x64' }
            0x14c   { 'x86' }
            default { ('0x{0:X}' -f $machine) }
        }
    } catch {
        'desconhecido'
    } finally {
        if ($br) { $br.Dispose() }
        if ($fs) { $fs.Dispose() }
    }
}

# Usado so para a mensagem amigavel "baixando versao X" — o $DownloadUrl em si
# e' version-agnostic (aponta sempre para "latest").
function Get-LatestReleaseTag {
    $apiUrl = "https://api.github.com/repos/$GithubOwner/$GithubRepo/releases/latest"
    $release = Invoke-RestMethod -Uri $apiUrl -Headers @{ 'User-Agent' = 'FinanceFmtTools-Install' } -ErrorAction Stop
    return $release.tag_name
}

# Dado uma raiz, encontra a pasta que contem $DllName: a propria raiz, uma
# subpasta 'bin\', ou (recursivamente) qualquer nivel abaixo. Retorna $null se nao achar.
function Find-BinDir {
    param([string]$Root)
    if (-not (Test-Path $Root)) { return $null }
    if (Test-Path (Join-Path $Root $DllName)) { return (Resolve-Path $Root).Path }
    $directBin = Join-Path $Root 'bin'
    if (Test-Path (Join-Path $directBin $DllName)) { return (Resolve-Path $directBin).Path }
    # Prioriza matches sob "\bin\" quando ha mais de um (ex.: build SDK-style tambem
    # deixa uma cópia em obj\) — evita instalar um artefato intermediario obsoleto.
    $hit = Get-ChildItem -Path $Root -Recurse -Filter $DllName -File -ErrorAction SilentlyContinue |
        Sort-Object { $_.FullName -notmatch '\\bin\\' } |
        Select-Object -First 1
    if ($hit) { return $hit.DirectoryName }
    return $null
}

# Remove a pasta temporaria de extracao, se foi criada (idempotente/silencioso).
function Remove-TempExtract {
    if ($script:TempExtractDir -and (Test-Path $script:TempExtractDir)) {
        Remove-Item -LiteralPath $script:TempExtractDir -Recurse -Force -ErrorAction SilentlyContinue
        $script:TempExtractDir = $null
    }
}

Write-Host ''
Write-Host '############################################################' -ForegroundColor White
Write-Host '#  Instalador - Finance Fmt Tools (C#)  v1.0.0            #' -ForegroundColor White
Write-Host '#  Sem admin | HKCU | sem regasm                          #' -ForegroundColor White
Write-Host '############################################################' -ForegroundColor White

# ===========================================================================
# PASSO 0 - Pre-instalacao: Excel fechado + checagem informativa de bitness
# ===========================================================================
Write-Step 'Pré-instalação'

Assert-ExcelNotRunning

Write-Step 'Detectando instalação VBA legada'
Remove-LegacyVbaAddin

# Bitness do Office — apenas informativo (Architectural Responsibility Map:
# "Bitness detection/guard (64-bit-only scope) — Installer script"). NUNCA
# bloqueia a instalacao; o add-in e AnyCPU e deve carregar mesmo em 32-bit,
# mas o baseline validado (CLAUDE.md/REQUIREMENTS.md) e Excel 64-bit.
$excelAppPathKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE'
if (Test-Path $excelAppPathKey) {
    $excelExePath = (Get-ItemProperty $excelAppPathKey).'(default)'
    if ($excelExePath -and (Test-Path $excelExePath)) {
        $excelArch = Test-PeMachine -Path $excelExePath
        Write-Info ("Excel detectado: {0} ({1})" -f $excelExePath, $excelArch)
        if ($excelArch -ne 'x64') {
            Write-Warn2 ("Excel parece ser {0} (baseline validado é x64). O add-in é AnyCPU e deve carregar mesmo assim, mas isto foge do cenário testado (ver FUT-01)." -f $excelArch)
        }
    }
}

$c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if (Test-Path $c2r) {
    $plat = (Get-ItemProperty $c2r).Platform
    if ($plat) {
        if ($plat -ne 'x64') {
            Write-Warn2 ("Office bitness (Click-to-Run): {0} — baseline validado é x64 (FUT-01)." -f $plat)
        } else {
            Write-Info ("Office bitness (Click-to-Run): {0}" -f $plat)
        }
    }
}

# ===========================================================================
# PASSO 1 - Resolver a pasta de origem dos binarios
# ===========================================================================
Write-Step 'Localizando os binários a instalar'

$SourceDir = $null

if ($Package) {
    # ----- Escotilha de teste local: -Package <zip> -----
    if (-not (Test-Path -LiteralPath $Package)) {
        Write-Err2 ("Pacote não encontrado: {0}" -f $Package)
        exit 1
    }
    $pkgFull = (Resolve-Path -LiteralPath $Package).Path
    if ([System.IO.Path]::GetExtension($pkgFull).ToLowerInvariant() -ne '.zip') {
        Write-Err2 ("O -Package deve ser um arquivo .zip. Recebido: {0}" -f $pkgFull)
        exit 1
    }
    Write-Info ("Pacote: {0}" -f $pkgFull)

    $script:TempExtractDir = Join-Path ([System.IO.Path]::GetTempPath()) ('financefmttools-install-{0}' -f ([System.Guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $script:TempExtractDir -Force | Out-Null
    try {
        Write-Info ("Extraindo para pasta temporária: {0}" -f $script:TempExtractDir)
        Expand-Archive -LiteralPath $pkgFull -DestinationPath $script:TempExtractDir -Force
    } catch {
        Write-Err2 ("Falha ao extrair o .zip: {0}" -f $_.Exception.Message)
        Remove-TempExtract
        exit 1
    }

    $SourceDir = Find-BinDir -Root $script:TempExtractDir
    if (-not $SourceDir) {
        Write-Err2 ("Não encontrei {0} dentro do .zip extraído." -f $DllName)
        Remove-TempExtract
        exit 1
    }
    Write-Ok ("Binários localizados no pacote extraído: {0}" -f $SourceDir)
}
elseif ($Source) {
    # ----- Escotilha de teste local: -Source <pasta> -----
    $SourceDir = Find-BinDir -Root $Source
    if (-not $SourceDir) {
        Write-Err2 ("Não encontrei {0} em -Source: {1}" -f $DllName, $Source)
        exit 1
    }
    Write-Ok ("Origem dos binários (-Source): {0}" -f $SourceDir)
}
else {
    # ----- Fluxo DOCUMENTADO (INST-01): baixar do GitHub Releases -----
    try {
        $tag = Get-LatestReleaseTag
        Write-Info ("Baixando versão {0}..." -f $tag)
    } catch {
        Write-Warn2 ("Não foi possível consultar a versão mais recente ({0}) — prosseguindo mesmo assim." -f $_.Exception.Message)
    }

    $script:TempExtractDir = Join-Path ([System.IO.Path]::GetTempPath()) ('financefmttools-install-{0}' -f ([System.Guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $script:TempExtractDir -Force | Out-Null
    $zipPath = Join-Path ([System.IO.Path]::GetTempPath()) ('financefmttools-{0}.zip' -f ([System.Guid]::NewGuid().ToString('N')))

    try {
        Write-Info ("Baixando: {0}" -f $DownloadUrl)
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $zipPath -UseBasicParsing
        Write-Info ("Extraindo para pasta temporária: {0}" -f $script:TempExtractDir)
        Expand-Archive -LiteralPath $zipPath -DestinationPath $script:TempExtractDir -Force
    } catch {
        Write-Err2 ("Falha ao baixar/extrair o release: {0}" -f $_.Exception.Message)
        Remove-TempExtract
        exit 1
    } finally {
        if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
    }

    $SourceDir = Find-BinDir -Root $script:TempExtractDir
    if (-not $SourceDir) {
        Write-Err2 ("Não encontrei {0} dentro do release baixado." -f $DllName)
        Remove-TempExtract
        exit 1
    }
    Write-Ok ("Binários baixados e extraídos: {0}" -f $SourceDir)
}

# Confere que todos os arquivos necessários existem na origem resolvida.
$missing = @()
foreach ($f in $AllFiles) {
    if (-not (Test-Path (Join-Path $SourceDir $f))) { $missing += $f }
}
if ($missing.Count -gt 0) {
    Write-Err2 ("Arquivos ausentes na origem: {0}" -f ($missing -join ', '))
    Remove-TempExtract
    exit 1
}
Write-Ok ("Encontrados os {0} binários: {1}" -f $AllFiles.Count, ($AllFiles -join ', '))

# --- Registro em HKCU continua abaixo (Task 2) ---

# ===========================================================================
# PASSO 2 - Instalação (cópia de arquivos + registro HKCU)
# ===========================================================================
Write-Step 'Instalação'

$valOk = $true

try {
    # Reconfere que o Excel continua fechado imediatamente antes de tocar em
    # arquivos — o download/extração acima pode ter levado tempo suficiente
    # para o usuário reabrir o Excel desde a checagem inicial (TOCTOU).
    Assert-ExcelNotRunning

    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    Write-Ok ("Pasta de instalação pronta: {0}" -f $InstallDir)

    foreach ($f in $AllFiles) {
        $src = Join-Path $SourceDir $f
        $dst = Join-Path $InstallDir $f
        Copy-Item -LiteralPath $src -Destination $dst -Force
        Write-Ok ("Copiado: {0}" -f $f)
    }

    $dllPath  = Join-Path $InstallDir $DllName
    # [Uri]::AbsoluteUri percent-codifica caracteres especiais (ex.: espaços em
    # "C:\Users\Nome Completo\...", comuns em máquinas corporativas) — uma
    # concatenação manual de string deixaria o CodeBase inválido nesses casos.
    $codeBase = ([Uri]$dllPath).AbsoluteUri
    # $AssemblyStr é lido do DLL recém-copiado, não hardcoded, para nunca ficar
    # dessincronizado de um bump de versão do assembly.
    $AssemblyStr = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath).FullName

    Write-Step 'Registrando em HKCU (sem admin)'

    # --- (a) Classe COM (shim CLR) em HKCU\Software\Classes -----------------
    $kProg      = "HKCU:\Software\Classes\$ProgId"
    $kProgClsid = "HKCU:\Software\Classes\$ProgId\CLSID"
    $kClsid     = "HKCU:\Software\Classes\CLSID\$Guid"
    $kClsidProg = "HKCU:\Software\Classes\CLSID\$Guid\ProgId"
    $kInproc    = "HKCU:\Software\Classes\CLSID\$Guid\InprocServer32"

    New-Item -Path $kProg      -Force | Out-Null
    New-Item -Path $kProgClsid -Force | Out-Null
    New-Item -Path $kClsid     -Force | Out-Null
    New-Item -Path $kClsidProg -Force | Out-Null
    New-Item -Path $kInproc    -Force | Out-Null

    Set-ItemProperty -Path $kProg      -Name '(default)' -Value $ProgId
    Set-ItemProperty -Path $kProgClsid -Name '(default)' -Value $Guid
    Set-ItemProperty -Path $kClsid     -Name '(default)' -Value $ClassName
    Set-ItemProperty -Path $kClsidProg -Name '(default)' -Value $ProgId

    Set-ItemProperty -Path $kInproc -Name '(default)'      -Value $Shim
    Set-ItemProperty -Path $kInproc -Name 'ThreadingModel' -Value $ThreadingMdl
    Set-ItemProperty -Path $kInproc -Name 'Class'          -Value $ClassName
    Set-ItemProperty -Path $kInproc -Name 'Assembly'       -Value $AssemblyStr
    Set-ItemProperty -Path $kInproc -Name 'RuntimeVersion' -Value $RuntimeVer
    Set-ItemProperty -Path $kInproc -Name 'CodeBase'       -Value $codeBase
    Write-Ok 'Classe COM registrada (CLSID + ProgId + InprocServer32).'

    # --- (b) Descoberta pelo Excel (NÃO versionado) -------------------------
    $kAddin = "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgId"
    New-Item -Path $kAddin -Force | Out-Null
    Set-ItemProperty -Path $kAddin -Name 'FriendlyName'  -Value $FriendlyName
    Set-ItemProperty -Path $kAddin -Name 'Description'   -Value $Description
    Set-ItemProperty -Path $kAddin -Name 'LoadBehavior'  -Value 3 -Type DWord
    Write-Ok 'Chave de add-in criada (LoadBehavior=3).'

    # --- (c) Resiliência anti-soft-disable (INST-03, versionado) ------------
    $kResil = "HKCU:\Software\Microsoft\Office\$OfficeVerKey\Excel\Resiliency\DoNotDisableAddinList"
    New-Item -Path $kResil -Force | Out-Null
    Set-ItemProperty -Path $kResil -Name $ProgId -Value 1 -Type DWord
    Write-Ok 'Chave de resiliência (DoNotDisableAddinList) criada.'
} catch {
    Write-Err2 ("Falha durante a instalação (cópia de arquivos ou registro): {0}" -f $_.Exception.Message)
    exit 1
} finally {
    Remove-TempExtract
}

# ===========================================================================
# PASSO 3 - Validação pós-instalação
# ===========================================================================
Write-Step 'Validação pós-instalação'

foreach ($f in $AllFiles) {
    if (Test-Path (Join-Path $InstallDir $f)) {
        Write-Ok ("Arquivo presente: {0}" -f $f)
    } else {
        Write-Err2 ("Arquivo NÃO copiado: {0}" -f $f); $valOk = $false
    }
}

$lb = (Get-ItemProperty -Path $kAddin -Name 'LoadBehavior' -ErrorAction SilentlyContinue).LoadBehavior
if ($lb -eq 3) { Write-Ok 'LoadBehavior = 3 (carregar no início).' }
else { Write-Err2 ("LoadBehavior inesperado: {0} (esperado 3)." -f $lb); $valOk = $false }

$cb = (Get-ItemProperty -Path $kInproc -Name 'CodeBase' -ErrorAction SilentlyContinue).CodeBase
if ($cb -eq $codeBase) { Write-Ok ("CodeBase correto: {0}" -f $cb) }
else { Write-Err2 ("CodeBase divergente: {0}" -f $cb); $valOk = $false }

# ===========================================================================
# PASSO 4 - Relatório final
# ===========================================================================
Write-Step 'Relatório final'

if ($valOk) { Write-Ok 'Instalação concluída com sucesso.' }
else { Write-Err2 'Instalação concluída COM PENDÊNCIAS (ver itens [ERRO] acima).' }

Write-Host ''
Write-Host 'O que foi instalado:' -ForegroundColor White
Write-Host ("  - Add-in COM '{0}' (CLSID {1})" -f $FriendlyName, $Guid)
Write-Host ("  - Binários em: {0}" -f $InstallDir)
Write-Host ("      {0}" -f ($AllFiles -join ', '))
Write-Host '  - Chaves de registro (HKCU, sem admin):'
Write-Host ("      HKCU\Software\Classes\CLSID\{0}\InprocServer32" -f $Guid)
Write-Host ("      HKCU\Software\Classes\{0}" -f $ProgId)
Write-Host ("      HKCU\Software\Microsoft\Office\Excel\Addins\{0}  (LoadBehavior=3)" -f $ProgId)
Write-Host ("      HKCU\Software\Microsoft\Office\{0}\Excel\Resiliency\DoNotDisableAddinList\{1}=1" -f $OfficeVerKey, $ProgId)

if ($script:VbaRemoved) {
    Write-Host ''
    Write-Host 'Migração automática:' -ForegroundColor White
    Write-Host ("  - Instalação legada VBA (.xlam) detectada e removida: {0}" -f $VbaXlamPath)
}

Write-Host ''
Write-Host 'Próximos passos:' -ForegroundColor White
Write-Host '  1. Abra o Excel.'
Write-Host '  2. Procure a aba "Finance Fmt" na Ribbon.'
Write-Host '  3. Para remover: rode scripts\uninstall.ps1.'
Write-Host ''

if ($valOk) { exit 0 } else { exit 1 }
