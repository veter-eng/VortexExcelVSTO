# Script para verificar instalação do VSTO Runtime em detalhes
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Verificando VSTO Runtime - Detalhado" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$found = $false

# Verificar várias versões e locais do VSTO Runtime
$paths = @(
    "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
    "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4",
    "HKLM:\SOFTWARE\Microsoft\VSTO_4.0",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO_4.0"
)

Write-Host "Procurando VSTO Runtime no Registry..." -ForegroundColor Yellow

foreach ($path in $paths) {
    if (Test-Path $path) {
        Write-Host "`n[ENCONTRADO] $path" -ForegroundColor Green
        $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        if ($props) {
            $props | Format-List
            $found = $true
        }
    }
}

if (-not $found) {
    Write-Host "`n[NÃO ENCONTRADO] VSTO Runtime não está instalado" -ForegroundColor Red
}

# Verificar arquivos do VSTO Runtime
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Verificando arquivos do VSTO Runtime" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$vstoFiles = @(
    "$env:ProgramFiles\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
    "$env:ProgramFiles\Common Files\Microsoft Shared\VSTO\10.0\VSTOLoader.dll",
    "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
    "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VSTO\10.0\VSTOLoader.dll"
)

$filesFound = $false
foreach ($file in $vstoFiles) {
    if (Test-Path $file) {
        Write-Host "[OK] $file" -ForegroundColor Green
        $filesFound = $true
    }
}

if (-not $filesFound) {
    Write-Host "[NÃO ENCONTRADO] Arquivos do VSTO Runtime não encontrados" -ForegroundColor Red
}

# Verificar Visual Studio
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Verificando Visual Studio" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$vsInstalled = $false
$vsPaths = @(
    "$env:ProgramFiles\Microsoft Visual Studio",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio"
)

foreach ($vsPath in $vsPaths) {
    if (Test-Path $vsPath) {
        Write-Host "[OK] Visual Studio encontrado em: $vsPath" -ForegroundColor Green
        $vsInstalled = $true

        # Procurar VSTO dentro do VS
        $vsVsto = Get-ChildItem -Path $vsPath -Recurse -Filter "VSTOInstaller.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($vsVsto) {
            Write-Host "[OK] VSTOInstaller.exe encontrado em: $($vsVsto.FullName)" -ForegroundColor Green
        }
    }
}

if (-not $vsInstalled) {
    Write-Host "[INFO] Visual Studio não está instalado" -ForegroundColor Yellow
}

# Resumo
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RESUMO" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

if ($found -or $filesFound) {
    Write-Host "[CONCLUSÃO] VSTO Runtime ESTÁ instalado!" -ForegroundColor Green
    Write-Host "O problema pode ser outra coisa." -ForegroundColor Yellow
} else {
    Write-Host "[CONCLUSÃO] VSTO Runtime NÃO está instalado!" -ForegroundColor Red
    Write-Host "`nOpções para instalar:" -ForegroundColor Yellow
    Write-Host "1. Instalar via Visual Studio Installer (se tiver VS instalado)" -ForegroundColor White
    Write-Host "2. Baixar VSTO Runtime standalone:" -ForegroundColor White
    Write-Host "   - https://aka.ms/vs/17/release/vstor_redist.exe" -ForegroundColor Cyan
    Write-Host "3. Instalar via winget:" -ForegroundColor White
    Write-Host "   winget install Microsoft.VSTORuntime" -ForegroundColor Cyan
}

Write-Host "`nPressione qualquer tecla para continuar..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
