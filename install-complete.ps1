# ========================================
# Vortex Excel Add-In - Instalação Completa Automatizada
# ========================================
# Este script automatiza todo o processo de instalação do plugin VSTO
# conforme descrito no README.md
# ========================================

param(
    [switch]$SkipPrerequisites,
    [switch]$SkipBuild,
    [switch]$SkipUninstall,
    [switch]$NoExcel
)

# Configuração de cores e formatação
$ErrorActionPreference = "Stop"
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
$script:totalSteps = 10

# Banner
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Vortex Excel Add-In" -ForegroundColor Cyan
Write-Host "Instalação Completa Automatizada" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Obter caminho do script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectPath = Join-Path $scriptPath "VortexExcelAddIn"
$binPath = Join-Path $projectPath "bin\Release"
$vstoFile = Join-Path $binPath "VortexExcelAddIn.vsto"

# ========================================
# PASSO 1: Verificar pré-requisitos
# ========================================
if (-not $SkipPrerequisites) {
    Write-Step "Verificando pré-requisitos..." "Cyan"
    
    # Verificar .NET Framework 4.8
    Write-Info "Verificando .NET Framework 4.8..."
    $netFramework = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Release -ErrorAction SilentlyContinue
    if ($netFramework -and $netFramework.Release -ge 528040) {
        Write-Success ".NET Framework 4.8 ou superior encontrado"
    } else {
        Write-Error ".NET Framework 4.8 não encontrado!"
        Write-Host ""
        Write-Host "Por favor, instale o .NET Framework 4.8:" -ForegroundColor Yellow
        Write-Host "https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Cyan
        Write-Host ""
        if (-not (Read-Host "Deseja continuar mesmo assim? (S/N)").StartsWith("S", "CurrentCultureIgnoreCase")) {
            exit 1
        }
    }
    
    # Verificar VSTO Runtime
    Write-Info "Verificando VSTO Runtime..."
    $vstoInstalled = $false
    $vstoPaths = @(
        "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
        "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4"
    )
    
    foreach ($path in $vstoPaths) {
        if (Test-Path $path) {
            $vstoInstalled = $true
            break
        }
    }
    
    # Verificar também arquivos físicos
    $vstoFiles = @(
        "$env:ProgramFiles\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
        "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
    )
    
    foreach ($file in $vstoFiles) {
        if (Test-Path $file) {
            $vstoInstalled = $true
            break
        }
    }
    
    if ($vstoInstalled) {
        Write-Success "VSTO Runtime encontrado"
    } else {
        Write-Error "VSTO Runtime não encontrado!"
        Write-Host ""
        Write-Host "Opções para instalar:" -ForegroundColor Yellow
        Write-Host "1. Execute: .\install-vsto.ps1" -ForegroundColor Cyan
        Write-Host "2. Baixe manualmente: https://aka.ms/vs/17/release/vstor_redist.exe" -ForegroundColor Cyan
        Write-Host "3. Via winget: winget install Microsoft.VSTORuntime" -ForegroundColor Cyan
        Write-Host ""
        if (-not (Read-Host "Deseja continuar mesmo assim? (S/N)").StartsWith("S", "CurrentCultureIgnoreCase")) {
            exit 1
        }
    }
    
    Write-Host ""
} else {
    Write-Step "Pulando verificação de pré-requisitos..." "Yellow"
    Write-Host ""
}

# ========================================
# PASSO 2: Fechar Excel
# ========================================
Write-Step "Fechando Excel se estiver aberto..." "Yellow"
$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Info "Encontrados $($excelProcesses.Count) processo(s) do Excel"
    Write-Info "Fechando processos..."
    $excelProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    # Verificar se fechou
    $stillRunning = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if ($stillRunning) {
        Write-Warning "Alguns processos do Excel ainda estão em execução"
        Write-Info "Tentando novamente..."
        Start-Sleep -Seconds 2
        $stillRunning | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    
    Write-Success "Excel fechado"
} else {
    Write-Success "Excel não está em execução"
}
Write-Host ""

# ========================================
# PASSO 3: Limpar itens desabilitados no registro
# ========================================
Write-Step "Limpando itens desabilitados no registro..." "Yellow"
$disabledItemsPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"
if (Test-Path $disabledItemsPath) {
    try {
        Remove-Item -Path $disabledItemsPath -Recurse -Force -ErrorAction Stop
        Write-Success "Itens desabilitados removidos"
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Warning "Não foi possível remover itens desabilitados: $errorMsg"
    }
} else {
    Write-Success "Nenhum item desabilitado encontrado"
}
Write-Host ""

# ========================================
# PASSO 4: Limpar cache de add-ins
# ========================================
Write-Step "Limpando cache de add-ins..." "Yellow"
$cachePaths = @(
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
    "$env:LOCALAPPDATA\Apps\2.0"
)

foreach ($cachePath in $cachePaths) {
    if (Test-Path $cachePath) {
        try {
            if ($cachePath -like "*Wef*") {
                # Limpar apenas arquivos, não pastas
                Get-ChildItem -Path $cachePath -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            } else {
                # Procurar por pastas relacionadas ao VortexExcel
                Get-ChildItem -Path $cachePath -Directory -ErrorAction SilentlyContinue | 
                    Where-Object { (Get-ChildItem $_.FullName -Recurse -ErrorAction SilentlyContinue | Select-String -Pattern "VortexExcel" -Quiet) } |
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-Success "Cache limpo: $(Split-Path $cachePath -Leaf)"
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Warning "Não foi possível limpar cache em $cachePath : $errorMsg"
        }
    }
}
Write-Host ""

# ========================================
# PASSO 5: Compilar o projeto
# ========================================
if (-not $SkipBuild) {
    Write-Step "Compilando o projeto..." "Yellow"
    
    # Procurar MSBuild
    $msbuildPaths = @(
        "C:\Program Files\Microsoft Visual Studio\2024\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2024\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2024\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSBuild.exe"
    )
    
    # Tambem procurar por versao numerica (18 = 2024)
    $vs18Paths = Get-ChildItem "C:\Program Files*\Microsoft Visual Studio\*\*\MSBuild\*\Bin\MSBuild.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($vs18Paths) {
        $msbuildPaths = @($vs18Paths) + $msbuildPaths
    }
    
    $msbuild = $null
    foreach ($path in $msbuildPaths) {
        if (Test-Path $path) {
            $msbuild = $path
            break
        }
    }
    
    # Tentar via PATH também
    if (-not $msbuild) {
        $msbuild = Get-Command msbuild -ErrorAction SilentlyContinue
        if ($msbuild) {
            $msbuild = $msbuild.Source
        }
    }
    
    if ($msbuild) {
        Write-Info "Usando MSBuild: $msbuild"
        $projectFile = Join-Path $projectPath "VortexExcelAddIn.csproj"
        
        if (Test-Path $projectFile) {
            # Restaurar pacotes NuGet primeiro
            Write-Info "Restaurando pacotes NuGet..."
            $packagesConfig = Join-Path $projectPath "packages.config"
            if (Test-Path $packagesConfig) {
                # Tentar restaurar via MSBuild primeiro
                & $msbuild $projectFile /t:Restore /v:minimal /nologo | Out-Null
                
                # Se falhar, tentar via NuGet.exe
                if ($LASTEXITCODE -ne 0) {
                    $nugetPath = "$env:TEMP\nuget.exe"
                    if (-not (Test-Path $nugetPath)) {
                        try {
                            Write-Info "Baixando NuGet.exe..."
                            Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $nugetPath -ErrorAction Stop
                        } catch {
                            Write-Warning "Nao foi possivel baixar NuGet.exe. Tentando continuar..."
                        }
                    }
                    
                    if (Test-Path $nugetPath) {
                        $packagesDir = Join-Path $scriptPath "packages"
                        & $nugetPath restore $packagesConfig -PackagesDirectory $packagesDir -NonInteractive | Out-Null
                    }
                }
            }
            
            Write-Info "Compilando em modo Release..."
            & $msbuild $projectFile /p:Configuration=Release /t:Clean,Build /v:minimal /nologo
            
            if ($LASTEXITCODE -eq 0) {
                Write-Success "Projeto compilado com sucesso"
            } else {
                Write-Error "Falha na compilação (código: $LASTEXITCODE)"
                Write-Host ""
                Write-Host "Tentando continuar com versão existente..." -ForegroundColor Yellow
            }
        } else {
            Write-Error "Arquivo de projeto não encontrado: $projectFile"
        }
    } else {
        Write-Warning "MSBuild não encontrado. Pulando compilação."
        Write-Info "Certifique-se de que o projeto já está compilado em Release"
    }
} else {
    Write-Step "Pulando compilação..." "Yellow"
    Write-Info "Usando versão existente"
}
Write-Host ""

# ========================================
# PASSO 6: Verificar se arquivo .vsto existe
# ========================================
Write-Step "Verificando arquivo de instalação..." "Yellow"
if (-not (Test-Path $vstoFile)) {
    Write-Error "Arquivo de instalação não encontrado: $vstoFile"
    Write-Host ""
    Write-Host "O arquivo .vsto não foi encontrado. Possíveis causas:" -ForegroundColor Yellow
    Write-Host "1. O projeto não foi compilado ainda" -ForegroundColor White
    Write-Host "2. A compilação falhou" -ForegroundColor White
    Write-Host "3. O caminho está incorreto" -ForegroundColor White
    Write-Host ""
    Write-Host "Tente compilar manualmente:" -ForegroundColor Yellow
    Write-Host "msbuild VortexExcelAddIn\VortexExcelAddIn.csproj /p:Configuration=Release" -ForegroundColor Cyan
    Write-Host ""
    exit 1
}
Write-Success "Arquivo .vsto encontrado"
Write-Host ""

# ========================================
# PASSO 7: Desinstalar versão anterior
# ========================================
if (-not $SkipUninstall) {
    Write-Step "Desinstalando versão anterior (se existir)..." "Yellow"
    
    # Tentar desinstalar via VSTOInstaller
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
            Write-Success "Desinstalação concluída"
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Warning "Não foi possível desinstalar via VSTOInstaller: $errorMsg"
        }
    } else {
        Write-Warning "VSTOInstaller não encontrado. Pulando desinstalação."
    }
} else {
    Write-Step "Pulando desinstalação..." "Yellow"
}
Write-Host ""

# ========================================
# PASSO 8: Configurar local confiável (opcional)
# ========================================
Write-Step "Configurando local confiável..." "Yellow"
try {
    $trustedLocationPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location99"
    
    if (-not (Test-Path $trustedLocationPath)) {
        New-Item -Path $trustedLocationPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $trustedLocationPath -Name "Path" -Value $binPath -Type String -Force | Out-Null
    Set-ItemProperty -Path $trustedLocationPath -Name "AllowSubFolders" -Value 1 -Type DWord -Force | Out-Null
    Set-ItemProperty -Path $trustedLocationPath -Name "Description" -Value "Vortex Excel Add-In" -Type String -Force | Out-Null
    
    Write-Success "Local confiável configurado"
} catch {
    $errorMsg = $_.Exception.Message
    Write-Warning "Não foi possível configurar local confiável: $errorMsg"
}
Write-Host ""

# ========================================
# PASSO 9: Instalar nova versão
# ========================================
Write-Step "Instalando nova versão..." "Yellow"
Write-Info "Abrindo instalador VSTO..."
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ATENCAO: Uma janela de instalacao sera aberta" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Por favor:" -ForegroundColor White
$inst1 = "1. Clique em 'Instalar' na janela que abrir"
$inst2 = "2. Aguarde a instalação concluir"
$inst3 = "3. Feche a janela de instalação quando terminar"
Write-Host $inst1 -ForegroundColor White
Write-Host $inst2 -ForegroundColor White
Write-Host $inst3 -ForegroundColor White
Write-Host ""

# Aguardar confirmação do usuário
$continue = Read-Host "Pressione ENTER quando estiver pronto para abrir o instalador"
Write-Host ""

try {
    Start-Process -FilePath $vstoFile -Wait
    Write-Success "Instalação concluída"
} catch {
    $errorMsg = $_.Exception.Message
    Write-Error "Erro ao abrir instalador: $errorMsg"
    Write-Host ""
    Write-Host "Tente abrir manualmente:" -ForegroundColor Yellow
    Write-Host $vstoFile -ForegroundColor Cyan
    exit 1
}
Write-Host ""

# ========================================
# PASSO 10: Abrir Excel e verificar
# ========================================
if (-not $NoExcel) {
    Write-Step "Abrindo Excel para verificar instalação..." "Yellow"
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "VERIFICACAO DA INSTALACAO" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Ao abrir o Excel, voce deve ver:" -ForegroundColor White
    Write-Host "[OK] 2 MessageBoxes de debug:" -ForegroundColor Green
    $msg1 = "  - 'Vortex Add-in: Iniciando...'"
    $msg2 = "  - 'Vortex Add-in: Carregado com sucesso!'"
    Write-Host $msg1 -ForegroundColor Gray
    Write-Host $msg2 -ForegroundColor Gray
    Write-Host ""
    Write-Host "[OK] Uma aba 'Vortex' no Ribbon" -ForegroundColor Green
    Write-Host "[OK] Um botao 'Vortex Plugin' dentro da aba" -ForegroundColor Green
    Write-Host ""
    Write-Host "Se NAO aparecerem essas mensagens:" -ForegroundColor Yellow
    Write-Host "1. Va em Arquivo > Opcoes > Suplementos" -ForegroundColor White
    $msg3 = "2. Selecione 'Itens Desabilitados' e clique em 'Ir...'"
    Write-Host $msg3 -ForegroundColor White
    $msg4 = "3. Se 'VortexExcelAddIn' estiver lá, selecione e clique em 'Habilitar'"
    Write-Host $msg4 -ForegroundColor White
    Write-Host "4. Execute: .\check-disabled.ps1 para verificar erros" -ForegroundColor White
    Write-Host ""
    
    $openExcel = Read-Host "Deseja abrir o Excel agora? (S/N)"
    if ($openExcel -match "^[Ss]") {
        Start-Process excel
        Write-Success "Excel aberto"
    } else {
        Write-Info "Excel não será aberto automaticamente"
    }
} else {
    Write-Step "Pulando abertura do Excel..." "Yellow"
}
Write-Host ""

# ========================================
# RESUMO FINAL
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "INSTALACAO CONCLUIDA!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Próximos passos:" -ForegroundColor Yellow
Write-Host "1. Abra o Microsoft Excel" -ForegroundColor White
$next1 = "2. Verifique se a aba 'Vortex' aparece no Ribbon"
$next2 = "3. Clique no botao 'Vortex Plugin' para abrir o painel"
$next3 = "4. Configure sua conexao na aba 'Configuracao'"
Write-Host $next1 -ForegroundColor White
Write-Host $next2 -ForegroundColor White
Write-Host $next3 -ForegroundColor White
Write-Host ""
Write-Host "Para mais informacoes, consulte o README.md" -ForegroundColor Cyan
Write-Host ""
Write-Host "Pressione qualquer tecla para sair..."
$null = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown)
