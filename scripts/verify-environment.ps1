<#
.SYNOPSIS
    Auditoria de ambiente (somente leitura) para o projeto Finance Fmt Tools (C#).

.DESCRIPTION
    Verifica os pre-requisitos do add-in COM "Finance Fmt Tools". Nao altera nada
    na maquina: apenas le e relata.

    Imprime um relatorio legivel com marcadores [OK] / [MISSING] / [FAIL] / [INFO] e
    retorna exit code 0 se todos os itens ESSENCIAIS estiverem presentes; caso
    contrario retorna exit code 1 (para permitir automacao).

    Idempotente e seguro: pode ser executado quantas vezes quiser.

    DOIS PERFIS DE PRE-REQUISITO (add-in COM em .NET Framework 4.8, instalado por
    copia + chaves HKCU):

      * ESSENCIAL PARA RODAR o add-in (maquina ALVO, usada por scripts\install.ps1):
          - .NET Framework 4.8
          - Excel (o add-in COM nao carrega em nenhuma variante que nao seja o
            Excel desktop classico)
        Sem isto o add-in nao carrega. Estes itens marcam falha global (exit 1).

      * ESSENCIAL PARA BUILDAR o add-in (maquina de DESENVOLVIMENTO):
          - .NET SDK ('dotnet'), Git, VS Code + extensao C#
        Sao necessarios para COMPILAR/manter o codigo, NAO para rodar o artefato.
        Por padrao tambem marcam falha global; com -RuntimeOnly o script verifica
        apenas o perfil de RUNTIME (util para checar a maquina alvo antes de instalar).

.PARAMETER RuntimeOnly
    Verifica apenas os itens ESSENCIAIS PARA RODAR o add-in (.NET Framework 4.8 +
    Excel). Itens de BUILD (SDK/Git/VS Code) viram informativos. Use na maquina
    ALVO (onde o add-in sera instalado), antes do install.ps1.

.NOTES
    Compativel com Windows PowerShell 5.1+ e PowerShell 7+.
#>

[CmdletBinding()]
param(
    [switch]$RuntimeOnly
)

$ErrorActionPreference = 'Continue'

# Quando -RuntimeOnly, itens de BUILD (SDK/Git/VS Code) NAO marcam falha global:
# tratamos a maquina como "alvo" (so precisa rodar o add-in, nao compila-lo).
$script:BuildEssential = -not $RuntimeOnly

# ---------------------------------------------------------------------------
# Infraestrutura de relatorio
# ---------------------------------------------------------------------------
$script:HasError = $false   # vira $true se algum item ESSENCIAL faltar/falhar

function Write-Status {
    param(
        [ValidateSet('OK', 'MISSING', 'FAIL', 'INFO')]
        [string]$Status,
        [string]$Message,
        [switch]$Essential   # se MISSING/FAIL num item essencial, marca falha global
    )
    $tag = '[{0}]' -f $Status
    $pad = $tag.PadRight(10)
    switch ($Status) {
        'OK'      { $color = 'Green' }
        'INFO'    { $color = 'Cyan' }
        'MISSING' { $color = 'Yellow' }
        'FAIL'    { $color = 'Red' }
    }
    Write-Host ($pad + $Message) -ForegroundColor $color
    if (($Status -eq 'MISSING' -or $Status -eq 'FAIL') -and $Essential) {
        $script:HasError = $true
    }
}

function Test-PeMachine {
    # Retorna 'x64', 'x86' ou uma string hex; le o cabecalho PE do executavel.
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
            0x8664 { return 'x64' }
            0x14c  { return 'x86' }
            default { return ('0x{0:X}' -f $machine) }
        }
    } catch {
        return 'desconhecido'
    } finally {
        if ($br) { $br.Dispose() }
        if ($fs) { $fs.Dispose() }
    }
}

Write-Host ''
Write-Host '=== verify-environment.ps1 ===' -ForegroundColor White
Write-Host ('Executado em: {0}' -f (Get-Date)) -ForegroundColor DarkGray
if ($RuntimeOnly) {
    Write-Host 'Modo: RUNTIME-ONLY (maquina ALVO: so o necessario para RODAR o add-in)' -ForegroundColor DarkGray
} else {
    Write-Host 'Modo: COMPLETO (BUILD + RUNTIME). Use -RuntimeOnly para checar so a maquina alvo.' -ForegroundColor DarkGray
}
Write-Host ''
Write-Host '----- ESSENCIAL PARA RODAR (maquina alvo): .NET Framework 4.8 + Excel -----' -ForegroundColor White

# ---------------------------------------------------------------------------
# 1. Windows
# ---------------------------------------------------------------------------
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    Write-Status -Status OK -Message ("Windows: {0} (build {1}, {2})" -f $os.Caption.Trim(), $os.BuildNumber, $os.OSArchitecture)
} catch {
    Write-Status -Status FAIL -Message "Windows: nao foi possivel detectar a versao." -Essential
}

# ---------------------------------------------------------------------------
# 2. Excel + bitness do Office
# ---------------------------------------------------------------------------
$excelExe = $null
$ap = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE'
if (Test-Path $ap) {
    $excelExe = (Get-ItemProperty $ap).'(default)'
}
if ($excelExe -and (Test-Path $excelExe)) {
    $ver = (Get-Item $excelExe).VersionInfo.FileVersion
    $arch = Test-PeMachine -Path $excelExe
    Write-Status -Status OK -Message ("Excel detectado: {0} ({1}) em {2}" -f $ver, $arch, $excelExe)
    if ($arch -ne 'x64') {
        Write-Status -Status INFO -Message ("Excel parece ser {0} (baseline validado e x64 - ver FUT-01)." -f $arch)
    }
} else {
    Write-Status -Status MISSING -Message "Excel (EXCEL.EXE) nao encontrado." -Essential
}

# Bitness do Office (Click-to-Run)
$c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if (Test-Path $c2r) {
    $cfg = Get-ItemProperty $c2r
    $plat = $cfg.Platform
    if ($plat) {
        Write-Status -Status OK -Message ("Office bitness: {0} -> o add-in deve ser {0} ou AnyCPU." -f $plat)
    } else {
        Write-Status -Status INFO -Message "Office bitness: nao reportado em ClickToRun\Configuration."
    }
} else {
    Write-Status -Status INFO -Message "Office ClickToRun\Configuration ausente (possivel instalacao MSI)."
}

Write-Host ''
Write-Host '----- ESSENCIAL PARA BUILDAR (maquina de desenvolvimento): .NET SDK, Git, VS Code -----' -ForegroundColor White
Write-Host '       (com -RuntimeOnly estes itens viram informativos, nao bloqueiam.)' -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 3. .NET SDK (ESSENCIAL PARA BUILDAR via 'dotnet' - NAO e necessario p/ rodar)
# ---------------------------------------------------------------------------
$dotnetCmd = Get-Command dotnet -ErrorAction SilentlyContinue
if ($dotnetCmd) {
    $sdks = & dotnet --list-sdks 2>$null
    if ($sdks) {
        $sdkList = ($sdks | ForEach-Object { ($_ -split ' ')[0] }) -join ', '
        Write-Status -Status OK -Message ("[BUILD] .NET SDK presente: {0} (dotnet em {1})" -f $sdkList, $dotnetCmd.Source)
    } else {
        Write-Status -Status MISSING -Message "[BUILD] 'dotnet' encontrado, mas NENHUM SDK instalado (so runtimes). Build via CLI nao funciona. Instale o .NET 8 SDK." -Essential:$script:BuildEssential
    }
} else {
    Write-Status -Status MISSING -Message "[BUILD] '.NET SDK' ausente. Instale: winget install --id Microsoft.DotNet.SDK.8 -e" -Essential:$script:BuildEssential
}

# ---------------------------------------------------------------------------
# 4. .NET Framework 4.8 (ESSENCIAL PARA RODAR o add-in - maquina alvo)
#    O add-in COM e net48; sem o .NET Framework 4.8 ele nao carrega no Excel.
# ---------------------------------------------------------------------------
$ndp = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
if (Test-Path $ndp) {
    $rel = (Get-ItemProperty $ndp).Release
    $fxVer = '4.x'
    if ($rel -ge 533320) { $fxVer = '4.8.1' }
    elseif ($rel -ge 528040) { $fxVer = '4.8' }
    if ($rel -ge 528040) {
        Write-Status -Status OK -Message ("[RUNTIME] .NET Framework {0} presente (Release {1}) - apto a RODAR o add-in." -f $fxVer, $rel)
    } else {
        Write-Status -Status MISSING -Message ("[RUNTIME] .NET Framework {0} (Release {1}) abaixo de 4.8. O add-in net48 exige 4.8+." -f $fxVer, $rel) -Essential
    }
} else {
    Write-Status -Status MISSING -Message "[RUNTIME] .NET Framework 4.8 NAO detectado. O add-in net48 NAO carregara sem ele." -Essential
}

# ---------------------------------------------------------------------------
# 5. MSBuild (Build Tools) - informativo
# ---------------------------------------------------------------------------
$msbuild = $null
$msbuildCmd = Get-Command msbuild -ErrorAction SilentlyContinue
if ($msbuildCmd) {
    $msbuild = $msbuildCmd.Source
} else {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $found = & $vswhere -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' 2>$null | Select-Object -First 1
        if ($found) { $msbuild = $found }
    }
}
if ($msbuild) {
    Write-Status -Status INFO -Message ("MSBuild disponivel: {0} (nao e usado pelo fluxo dotnet CLI deste projeto - ver CLAUDE.md)." -f $msbuild)
} else {
    Write-Status -Status INFO -Message "MSBuild (Build Tools) nao encontrado. Nao e necessario - este projeto usa 'dotnet build'."
}

# ---------------------------------------------------------------------------
# 6. PowerShell
# ---------------------------------------------------------------------------
$psv = $PSVersionTable.PSVersion
if ($psv.Major -ge 5) {
    Write-Status -Status OK -Message ("PowerShell: {0} ({1})" -f $psv.ToString(), $PSVersionTable.PSEdition) -Essential
} else {
    Write-Status -Status FAIL -Message ("PowerShell muito antigo: {0}. Requer 5.1+." -f $psv.ToString()) -Essential
}
$pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
if ($pwsh) {
    Write-Status -Status INFO -Message ("PowerShell 7+ tambem presente: {0}" -f $pwsh.Version.ToString())
}

# ---------------------------------------------------------------------------
# 7. Git
# ---------------------------------------------------------------------------
$git = Get-Command git -ErrorAction SilentlyContinue
if ($git) {
    Write-Status -Status OK -Message ("[BUILD] Git: {0}" -f ((& git --version) -replace 'git version ', ''))
} else {
    Write-Status -Status MISSING -Message "[BUILD] Git ausente. Instale: winget install --id Git.Git -e" -Essential:$script:BuildEssential
}

# ---------------------------------------------------------------------------
# 8. VS Code + extensao C#
# ---------------------------------------------------------------------------
$code = Get-Command code -ErrorAction SilentlyContinue
if ($code) {
    $cv = (& code --version 2>$null | Select-Object -First 1)
    $exts = & code --list-extensions 2>$null
    $hasCsharp = $exts -contains 'ms-dotnettools.csharp'
    if ($hasCsharp) {
        Write-Status -Status OK -Message ("[BUILD] VS Code: {0} (extensao C# 'ms-dotnettools.csharp' instalada)" -f $cv)
    } else {
        Write-Status -Status MISSING -Message ("[BUILD] VS Code {0} presente, mas SEM a extensao C#. Instale: code --install-extension ms-dotnettools.csharp" -f $cv) -Essential:$script:BuildEssential
    }
} else {
    Write-Status -Status MISSING -Message "[BUILD] VS Code ausente. Instale: winget install --id Microsoft.VisualStudioCode -e" -Essential:$script:BuildEssential
}

# ---------------------------------------------------------------------------
# 9. winget (auxiliar de instalacao)
# ---------------------------------------------------------------------------
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($winget) {
    Write-Status -Status INFO -Message ("winget disponivel: {0}" -f ((& winget --version) -join ''))
} else {
    Write-Status -Status INFO -Message "winget nao encontrado (instalacao manual de ferramentas pode ser necessaria)."
}

# ---------------------------------------------------------------------------
# Resumo / exit code
# ---------------------------------------------------------------------------
Write-Host ''
if ($script:HasError) {
    if ($RuntimeOnly) {
        Write-Host 'RESULTADO: faltam itens ESSENCIAIS PARA RODAR o add-in (.NET Framework 4.8 / Excel).' -ForegroundColor Red
    } else {
        Write-Host 'RESULTADO: ha itens ESSENCIAIS ausentes/falhos (ver [RUNTIME]/[BUILD] e [MISSING]/[FAIL] acima).' -ForegroundColor Red
    }
    exit 1
} else {
    if ($RuntimeOnly) {
        Write-Host 'RESULTADO: maquina APTA A RODAR o add-in (RUNTIME ok: .NET Framework 4.8 + Excel).' -ForegroundColor Green
    } else {
        Write-Host 'RESULTADO: ambiente apto (RUNTIME para rodar + BUILD para compilar - todos essenciais OK).' -ForegroundColor Green
    }
    exit 0
}
