# ========================================
# Vortex Excel Add-In - Limpeza Completa do Ambiente
# ========================================
# Este script remove completamente o plugin e limpa o ambiente
# para permitir uma instalação limpa do zero
# ========================================

param(
    [switch]$RemoveVSTORuntime,
    [switch]$SkipNuGetPackages,
    [switch]$SkipBinObj
)

# Configuração de cores e formatação
$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

function Write-Step {
    param([string]$Message, [string]$Color = "Yellow")
    Write-Host "[$($script:stepNumber)/$($script:totalSteps)] $Message" -ForegroundColor $Color
    $script:stepNumber++
}

function Write-Success {
    param([string]$Message)
    Write-Host "       [OK] $Message" -ForegroundColor Green
}

function Write-Error {
    param([string]$Message)
    Write-Host "       [ERRO] $Message" -ForegroundColor Red
}

function Write-Warning {
    param([string]$Message)
    Write-Host "       [AVISO] $Message" -ForegroundColor Yellow
}

function Write-Info {
    param([string]$Message)
    Write-Host "       [INFO] $Message" -ForegroundColor Cyan
}

# Inicializar contadores
$script:stepNumber = 1
$script:totalSteps = 8

# Banner
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Vortex Excel Add-In" -ForegroundColor White
Write-Host "Limpeza Completa do Ambiente" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Obter caminhos
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectPath = Join-Path $scriptPath "VortexExcelAddIn"
$vstoFile = Join-Path $projectPath "bin\Release\VortexExcelAddIn.vsto"

# Verificar possíveis caminhos do arquivo .vsto
$possibleVstoPaths = @(
    $vstoFile
    (Join-Path $projectPath "bin\Debug\VortexExcelAddIn.vsto")
)

$vstoFile = $null
foreach ($path in $possibleVstoPaths) {
    if (Test-Path $path) {
        $vstoFile = $path
        break
    }
}

# ========================================
# PASSO 1: Fechar Excel
# ========================================
Write-Step "Fechando Excel se estiver aberto..." "Yellow"
$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Info "Encontrados $($excelProcesses.Count) processo(s) do Excel"
    try {
        $excelProcesses | Stop-Process -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Write-Success "Excel fechado"
    } catch {
        Write-Warning "Nao foi possivel fechar o Excel: $($_.Exception.Message)"
    }
} else {
    Write-Success "Excel nao esta em execucao"
}
Write-Host ""

# ========================================
# PASSO 2: Desinstalar plugin VSTO
# ========================================
Write-Step "Desinstalando plugin VSTO..." "Yellow"
if ($vstoFile) {
    $vstoInstallerPaths = @(
        "$env:ProgramFiles\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
        "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
    )
    
    $vstoInstaller = $null
    foreach ($path in $vstoInstallerPaths) {
        if (Test-Path $path) {
            $vstoInstaller = $path
            break
        }
    }
    
    if ($vstoInstaller) {
        Write-Info "Desinstalando via VSTOInstaller..."
        try {
            $process = Start-Process -FilePath $vstoInstaller -ArgumentList "/uninstall", "`"$vstoFile`"", "/silent" -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Write-Success "Plugin desinstalado"
        } catch {
            Write-Warning "Nao foi possivel desinstalar via VSTOInstaller: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "VSTOInstaller nao encontrado"
    }
} else {
    Write-Info "Arquivo .vsto nao encontrado, pulando desinstalacao"
}
Write-Host ""

# ========================================
# PASSO 3: Limpar itens desabilitados no registro
# ========================================
Write-Step "Limpando itens desabilitados no registro..." "Yellow"
$disabledItemsPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems"
)

foreach ($disabledPath in $disabledItemsPaths) {
    if (Test-Path $disabledPath) {
        try {
            Remove-Item -Path $disabledPath -Recurse -Force -ErrorAction Stop
            Write-Success "Itens desabilitados removidos: $(Split-Path $disabledPath -Leaf)"
        } catch {
            Write-Warning "Nao foi possivel remover: $disabledPath"
        }
    }
}
Write-Host ""

# ========================================
# PASSO 4: Limpar cache de add-ins
# ========================================
Write-Step "Limpando cache de add-ins..." "Yellow"
$cachePaths = @(
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
    "$env:LOCALAPPDATA\Microsoft\Office\15.0\Wef",
    "$env:LOCALAPPDATA\Apps\2.0"
)

foreach ($cachePath in $cachePaths) {
    if (Test-Path $cachePath) {
        try {
            if ($cachePath -like "*Wef*") {
                # Limpar apenas arquivos, não pastas
                $files = Get-ChildItem -Path $cachePath -File -ErrorAction SilentlyContinue
                if ($files) {
                    $files | Remove-Item -Force -ErrorAction SilentlyContinue
                    Write-Success "Cache limpo: $(Split-Path $cachePath -Leaf)"
                }
            } else {
                # Procurar por pastas relacionadas ao VortexExcel
                $folders = Get-ChildItem -Path $cachePath -Directory -ErrorAction SilentlyContinue | 
                    Where-Object { 
                        $content = Get-ChildItem $_.FullName -Recurse -ErrorAction SilentlyContinue | Select-String -Pattern "VortexExcel" -Quiet
                        $content
                    }
                if ($folders) {
                    $folders | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Success "Pastas relacionadas ao Vortex removidas de: $(Split-Path $cachePath -Leaf)"
                }
            }
        } catch {
            Write-Warning "Nao foi possivel limpar cache em $cachePath : $($_.Exception.Message)"
        }
    }
}
Write-Host ""

# ========================================
# PASSO 5: Remover locais confiáveis do registro
# ========================================
Write-Step "Removendo locais confiáveis do registro..." "Yellow"
$trustedLocationPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location99",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Security\Trusted Locations\Location99"
)

foreach ($trustedPath in $trustedLocationPaths) {
    if (Test-Path $trustedPath) {
        try {
            $locationValue = (Get-ItemProperty -Path $trustedPath -Name "Path" -ErrorAction SilentlyContinue).Path
            if ($locationValue -and $locationValue -like "*Vortex*") {
                Remove-Item -Path $trustedPath -Recurse -Force -ErrorAction Stop
                Write-Success "Local confiavel removido: $locationValue"
            }
        } catch {
            Write-Warning "Nao foi possivel remover local confiavel: $trustedPath"
        }
    }
}
Write-Host ""

# ========================================
# PASSO 6: Remover pacotes NuGet (opcional)
# ========================================
if (-not $SkipNuGetPackages) {
    Write-Step "Removendo pacotes NuGet..." "Yellow"
    $packagesDir = Join-Path $scriptPath "packages"
    if (Test-Path $packagesDir) {
        try {
            Remove-Item -Path $packagesDir -Recurse -Force -ErrorAction Stop
            Write-Success "Pacotes NuGet removidos"
        } catch {
            Write-Warning "Nao foi possivel remover pacotes NuGet: $($_.Exception.Message)"
        }
    } else {
        Write-Success "Diretorio de pacotes nao encontrado"
    }
} else {
    Write-Step "Pulando remocao de pacotes NuGet..." "Yellow"
}
Write-Host ""

# ========================================
# PASSO 7: Remover pastas bin e obj (opcional)
# ========================================
if (-not $SkipBinObj) {
    Write-Step "Removendo pastas bin e obj..." "Yellow"
    $foldersToRemove = @(
        (Join-Path $projectPath "bin"),
        (Join-Path $projectPath "obj")
    )
    
    foreach ($folder in $foldersToRemove) {
        if (Test-Path $folder) {
            try {
                Remove-Item -Path $folder -Recurse -Force -ErrorAction Stop
                Write-Success "Pasta removida: $(Split-Path $folder -Leaf)"
            } catch {
                Write-Warning "Nao foi possivel remover: $folder"
            }
        }
    }
} else {
    Write-Step "Pulando remocao de pastas bin e obj..." "Yellow"
}
Write-Host ""

# ========================================
# PASSO 8: Remover VSTO Runtime (opcional)
# ========================================
if ($RemoveVSTORuntime) {
    Write-Step "Removendo VSTO Runtime..." "Yellow"
    
    # Tentar remover via winget
    $wingetFound = $false
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Info "Tentando remover via winget..."
        $wingetPackages = @(
            "Microsoft.VSTORuntime",
            "Microsoft.VSTORuntime.4.0"
        )
        
        foreach ($packageId in $wingetPackages) {
            try {
                $wingetOutput = & winget uninstall $packageId --silent --accept-source-agreements --accept-package-agreements 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Success "VSTO Runtime removido via winget: $packageId"
                    $wingetFound = $true
                    break
                }
            } catch {
                # Continuar tentando outros pacotes
            }
        }
    }
    
    if (-not $wingetFound) {
        Write-Warning "Nao foi possivel remover VSTO Runtime automaticamente"
        Write-Host ""
        Write-Host "Para remover manualmente:" -ForegroundColor Yellow
        Write-Host "1. Abra 'Adicionar ou remover programas'" -ForegroundColor White
        Write-Host "2. Procure por 'Microsoft Visual Studio 2010 Tools for Office Runtime'" -ForegroundColor White
        Write-Host "3. Clique em 'Desinstalar'" -ForegroundColor White
        Write-Host ""
    }
} else {
    Write-Step "Pulando remocao do VSTO Runtime..." "Yellow"
    Write-Info "Use -RemoveVSTORuntime para remover o VSTO Runtime tambem"
}
Write-Host ""

# ========================================
# Resumo final
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Limpeza concluida!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "O ambiente foi limpo com sucesso." -ForegroundColor White
Write-Host "Voce pode agora executar install-complete.ps1 para uma instalacao limpa." -ForegroundColor White
Write-Host ""
