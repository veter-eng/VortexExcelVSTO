# Script para instalar VSTO Runtime automaticamente
# Requires -RunAsAdministrator

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Instalador do VSTO Runtime" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# URLs alternativas para VSTO Runtime
$urls = @(
    "https://aka.ms/vs/17/release/vstor_redist.exe",
    "https://download.visualstudio.microsoft.com/download/pr/100349138/6ba2f067a0b6cb0547888cf3dacf904d/vstor_redist.exe"
)

$tempFile = "$env:TEMP\vstor_redist.exe"

# Tentar baixar de diferentes URLs
$downloaded = $false
foreach ($url in $urls) {
    Write-Host "Tentando baixar de: $url" -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $url -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
        if (Test-Path $tempFile) {
            $fileSize = (Get-Item $tempFile).Length
            if ($fileSize -gt 1MB) {
                Write-Host "[OK] Download concluído! ($([math]::Round($fileSize/1MB, 2)) MB)" -ForegroundColor Green
                $downloaded = $true
                break
            }
        }
    }
    catch {
        Write-Host "[ERRO] Falha ao baixar de $url" -ForegroundColor Red
        Write-Host "      $($_.Exception.Message)" -ForegroundColor Red
    }
}

if (-not $downloaded) {
    Write-Host "`n[ERRO] Não foi possível baixar o VSTO Runtime de nenhuma URL" -ForegroundColor Red
    Write-Host "`nTentando método alternativo via winget..." -ForegroundColor Yellow

    # Tentar via winget
    try {
        $wingetExists = Get-Command winget -ErrorAction SilentlyContinue
        if ($wingetExists) {
            Write-Host "Instalando via winget..." -ForegroundColor Yellow
            winget install --id Microsoft.VSTORuntime --accept-package-agreements --accept-source-agreements
            Write-Host "[OK] Instalação via winget concluída!" -ForegroundColor Green
            exit 0
        } else {
            Write-Host "[ERRO] winget não está disponível" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "[ERRO] Falha ao instalar via winget: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "INSTALACAO MANUAL NECESSARIA" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Nao foi possivel baixar automaticamente." -ForegroundColor White
    Write-Host ""
    Write-Host "Por favor, baixe manualmente de uma das opcoes:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "OPCAO 1: Visual Studio Community (Recomendado)" -ForegroundColor Cyan
    Write-Host "  - O Visual Studio inclui o VSTO Runtime automaticamente" -ForegroundColor White
    Write-Host "  - Download: https://visualstudio.microsoft.com/pt-br/thank-you-downloading-visual-studio/?sku=Community&channel=Stable&version=VS18&source=VSLandingPage&passive=false&cid=2500" -ForegroundColor White
    Write-Host "  - Durante a instalacao, selecione 'Desenvolvimento para Desktop com .NET'" -ForegroundColor White
    Write-Host ""
    Write-Host "OPCAO 2: Link direto do VSTO Runtime" -ForegroundColor Cyan
    Write-Host "  https://aka.ms/vs/17/release/vstor_redist.exe" -ForegroundColor White
    Write-Host ""
    Write-Host "OPCAO 3: Pagina oficial da Microsoft" -ForegroundColor Cyan
    Write-Host "  https://www.microsoft.com/en-us/download/details.aspx?id=105890" -ForegroundColor White
    Write-Host ""
    
    $openBrowser = Read-Host "Deseja abrir a pagina de download do Visual Studio Community? (S/N)"
    if (-not [string]::IsNullOrWhiteSpace($openBrowser) -and $openBrowser -match "^[Ss]") {
        Write-Host ""
        Write-Host "Abrindo pagina de download do Visual Studio Community..." -ForegroundColor Yellow
        Start-Process "https://visualstudio.microsoft.com/pt-br/thank-you-downloading-visual-studio/?sku=Community&channel=Stable&version=VS18&source=VSLandingPage&passive=false&cid=2500"
        Write-Host ""
        Write-Host "Apos instalar o Visual Studio, o VSTO Runtime sera instalado automaticamente." -ForegroundColor Yellow
        Write-Host "Depois execute este script novamente para verificar." -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Pressione qualquer tecla para sair..."
    $null = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown)
    exit 1
}

# Instalar
Write-Host "`nIniciando instalação do VSTO Runtime..." -ForegroundColor Yellow
try {
    $process = Start-Process -FilePath $tempFile -ArgumentList "/quiet /norestart" -Wait -PassThru

    if ($process.ExitCode -eq 0) {
        Write-Host "[OK] VSTO Runtime instalado com sucesso!" -ForegroundColor Green
    }
    elseif ($process.ExitCode -eq 3010) {
        Write-Host "[OK] VSTO Runtime instalado! (Reinicialização necessária)" -ForegroundColor Yellow
    }
    else {
        Write-Host "[AVISO] Instalação retornou código: $($process.ExitCode)" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "[ERRO] Falha ao instalar: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    # Limpar arquivo temporário
    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Instalação concluída!" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Proximos passos:" -ForegroundColor Yellow
Write-Host "1. Execute: .\install-complete.ps1 para instalar o plugin" -ForegroundColor White
Write-Host "2. Ou reinicie o computador primeiro (recomendado se solicitado)" -ForegroundColor White
Write-Host ""

Write-Host "Pressione qualquer tecla para sair..."
$null = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown)
