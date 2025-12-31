# Build and Install Script
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Vortex Excel Add-In - Build e Instalação" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if Excel is running
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    Write-Host "[ERRO] O Excel está em execução!" -ForegroundColor Red
    Write-Host "Por favor, feche o Excel e execute este script novamente." -ForegroundColor Yellow
    exit 1
}

# Clean cache
Write-Host "[1/4] Limpando cache de add-ins..." -ForegroundColor Yellow
Remove-Item "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*.*" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Apps\2.0\*VortexExcel*" -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "       Cache limpo!" -ForegroundColor Green
Write-Host ""

# Build
Write-Host "[2/4] Compilando o projeto..." -ForegroundColor Yellow
$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$project = "$PSScriptRoot\VortexExcelAddIn\VortexExcelAddIn.csproj"

& $msbuild $project /p:Configuration=Release /t:Clean,Build /v:minimal

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERRO] Falha na compilação!" -ForegroundColor Red
    exit 1
}
Write-Host "       Compilado com sucesso!" -ForegroundColor Green
Write-Host ""

# Uninstall previous version
Write-Host "[3/4] Desinstalando versão anterior (se existir)..." -ForegroundColor Yellow
$vsto = "$PSScriptRoot\VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto"
if (Test-Path $vsto) {
    Start-Process -FilePath $vsto -ArgumentList "/uninstall", "/silent" -Wait -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}
Write-Host "       Desinstalação concluída!" -ForegroundColor Green
Write-Host ""

# Install new version
Write-Host "[4/4] Instalando nova versão..." -ForegroundColor Yellow
if (Test-Path $vsto) {
    Start-Process -FilePath $vsto -Wait
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Instalação concluída!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Próximos passos:" -ForegroundColor Yellow
    Write-Host "1. Abra o Microsoft Excel"
    Write-Host "2. Procure pela aba 'Vortex'"
    Write-Host "3. Clique no botão 'Auto-Refresh' para configurar!"
    Write-Host ""
} else {
    Write-Host "[ERRO] Arquivo VSTO não encontrado!" -ForegroundColor Red
    exit 1
}
