# Script de Diagnóstico do Vortex Excel Plugin
# Execute este script no PowerShell como Administrador para verificar todos os requisitos

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Diagnóstico do Vortex Excel Plugin" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

$errors = 0
$warnings = 0

# 1. Verificar versão do Windows
Write-Host "[1/7] Verificando versão do Windows..." -ForegroundColor Yellow
$os = Get-WmiObject -Class Win32_OperatingSystem
Write-Host "  Sistema: $($os.Caption)" -ForegroundColor Gray
Write-Host "  Versão: $($os.Version)" -ForegroundColor Gray
Write-Host "  ✓ OK" -ForegroundColor Green
Write-Host ""

# 2. Verificar se o Excel está instalado
Write-Host "[2/7] Verificando instalação do Microsoft Excel..." -ForegroundColor Yellow
$excelInstalled = $false

# Tentar encontrar Excel.exe
$excelExePaths = @(
    "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
    "${env:ProgramFiles}\Microsoft Office\Office16\EXCEL.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE"
)

foreach ($path in $excelExePaths) {
    if (Test-Path $path) {
        Write-Host "  ✓ Excel encontrado: $path" -ForegroundColor Green
        $excelInstalled = $true
        break
    }
}

# Se não encontrou pelo caminho, tentar pelo registro
if (-not $excelInstalled) {
    $excelPath = "HKLM:\SOFTWARE\Microsoft\Office"
    if (Test-Path $excelPath) {
        $versions = Get-ChildItem $excelPath -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match "^\d+\.\d+$" }
        if ($versions) {
            foreach ($version in $versions) {
                $excelKey = Join-Path $excelPath "$($version.PSChildName)\Excel"
                if (Test-Path $excelKey) {
                    Write-Host "  ✓ Excel encontrado (versao $($version.PSChildName))" -ForegroundColor Green
                    $excelInstalled = $true
                    break
                }
            }
        }
    }
}

if (-not $excelInstalled) {
    Write-Host "  ✗ Excel nao encontrado" -ForegroundColor Red
    Write-Host "    O plugin esta registrado mas o Excel nao foi detectado" -ForegroundColor Yellow
    Write-Host "    Isso pode ser um falso positivo se o Excel estiver instalado via Microsoft 365" -ForegroundColor Yellow
    $warnings++
}
Write-Host ""

# 3. Verificar .NET Framework 4.8
Write-Host "[3/7] Verificando .NET Framework 4.8..." -ForegroundColor Yellow
try {
    $netfxPath = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    $release = (Get-ItemProperty $netfxPath -ErrorAction Stop).Release
    Write-Host "  Release: $release" -ForegroundColor Gray

    if ($release -ge 528040) {
        $version = switch ($release) {
            { $_ -ge 533320 } { "4.8.1 ou superior" }
            { $_ -ge 528040 } { "4.8" }
            default { "Inferior a 4.8" }
        }
        Write-Host "  ✓ .NET Framework $version instalado" -ForegroundColor Green
    } else {
        Write-Host "  ✗ .NET Framework 4.8 não encontrado (Release: $release)" -ForegroundColor Red
        Write-Host "    Baixe em: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Yellow
        $errors++
    }
} catch {
    Write-Host "  ✗ .NET Framework 4.8 não encontrado" -ForegroundColor Red
    Write-Host "    Baixe em: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Yellow
    $errors++
}
Write-Host ""

# 4. Verificar recursos opcionais do .NET
Write-Host "[4/7] Verificando recursos opcionais do .NET..." -ForegroundColor Yellow
$netfxFeatures = Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -like "*NetFx*"}
$requiredFeatures = @("NetFx3", "NetFx4-AdvSrvs")

foreach ($feature in $requiredFeatures) {
    $found = $netfxFeatures | Where-Object { $_.FeatureName -eq $feature }
    if ($found) {
        if ($found.State -eq "Enabled") {
            Write-Host "  ✓ $($found.FeatureName) - Habilitado" -ForegroundColor Green
        } else {
            Write-Host "  ✗ $($found.FeatureName) - Desabilitado" -ForegroundColor Red
            Write-Host "    Execute: Enable-WindowsOptionalFeature -Online -FeatureName $($found.FeatureName)" -ForegroundColor Yellow
            $errors++
        }
    } else {
        Write-Host "  ⚠ ${feature} - Nao encontrado" -ForegroundColor Yellow
        $warnings++
    }
}
Write-Host ""

# 5. Verificar assemblies WPF
Write-Host "[5/7] Verificando assemblies WPF..." -ForegroundColor Yellow
$wpfAssemblies = @(
    "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\WPF\PresentationFramework.dll",
    "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\WPF\PresentationCore.dll",
    "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\WPF\WindowsBase.dll",
    "$env:SystemRoot\assembly\GAC_MSIL\WindowsFormsIntegration\4.0.0.0__31bf3856ad364e35\WindowsFormsIntegration.dll"
)

foreach ($assembly in $wpfAssemblies) {
    $assemblyName = Split-Path $assembly -Leaf
    if (Test-Path $assembly) {
        Write-Host "  ✓ $assemblyName" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $assemblyName não encontrado" -ForegroundColor Red
        Write-Host "    Caminho: $assembly" -ForegroundColor Gray
        $errors++
    }
}
Write-Host ""

# 6. Verificar VSTO Runtime
Write-Host "[6/7] Verificando VSTO Runtime..." -ForegroundColor Yellow
$vstoFound = $false

# Verificar em múltiplos locais
$vstoPaths = @(
    "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R",
    "HKLM:\SOFTWARE\Microsoft\VSTO_Runtime\v4"
)

foreach ($vstoPath in $vstoPaths) {
    if (Test-Path $vstoPath) {
        $vstoVersion = (Get-ItemProperty $vstoPath -ErrorAction SilentlyContinue).Version
        if ($vstoVersion) {
            Write-Host "  ✓ VSTO Runtime instalado (Versao: $vstoVersion)" -ForegroundColor Green
            $vstoFound = $true
            break
        }
    }
}

# Verificar também por DLL específica
if (-not $vstoFound) {
    $vstoDll = "${env:ProgramFiles}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
    if (Test-Path $vstoDll) {
        Write-Host "  ✓ VSTO Runtime instalado (encontrado VSTOInstaller.exe)" -ForegroundColor Green
        $vstoFound = $true
    }
}

if (-not $vstoFound) {
    Write-Host "  ✗ VSTO Runtime nao encontrado" -ForegroundColor Red
    Write-Host "    Baixe em: https://www.microsoft.com/download/details.aspx?id=56961" -ForegroundColor Yellow
    Write-Host "    IMPORTANTE: Reinicie o computador apos instalar!" -ForegroundColor Yellow
    $errors++
}
Write-Host ""

# 7. Verificar certificado do plugin
Write-Host "[7/8] Verificando certificado do plugin..." -ForegroundColor Yellow
$certThumbprint = "4D194630B60D2B3CB959B37D35317EA72E7DA651"
$certFound = $false

# Verificar em Trusted Publishers
$trustedCert = Get-ChildItem -Path Cert:\CurrentUser\TrustedPublisher -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $certThumbprint }
if ($trustedCert) {
    Write-Host "  ✓ Certificado encontrado em Trusted Publishers" -ForegroundColor Green
    Write-Host "    Emissor: $($trustedCert.Issuer)" -ForegroundColor Gray
    Write-Host "    Valido ate: $($trustedCert.NotAfter)" -ForegroundColor Gray
    $certFound = $true
} else {
    Write-Host "  ✗ Certificado NAO encontrado em Trusted Publishers" -ForegroundColor Red
    Write-Host "    Thumbprint esperado: $certThumbprint" -ForegroundColor Gray
    Write-Host "" -ForegroundColor Yellow
    Write-Host "    ISTO E PROVAVELMENTE O PROBLEMA!" -ForegroundColor Red
    Write-Host "    O plugin pode aparecer no Ribbon mas a barra lateral nao abre" -ForegroundColor Yellow
    Write-Host "    sem o certificado confiavel." -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Yellow
    Write-Host "    Solucao: Execute o script install-with-certificate.bat" -ForegroundColor Yellow
    $errors++
}

# Verificar também em outras lojas de certificados
$rootCert = Get-ChildItem -Path Cert:\CurrentUser\Root -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $certThumbprint }
if ($rootCert) {
    Write-Host "  ℹ Certificado também encontrado em Root" -ForegroundColor Cyan
    $certFound = $true
}

if (-not $certFound) {
    Write-Host "" -ForegroundColor Yellow
    Write-Host "  Certificados disponiveis em Trusted Publishers:" -ForegroundColor Gray
    $allTrustedCerts = Get-ChildItem -Path Cert:\CurrentUser\TrustedPublisher -ErrorAction SilentlyContinue
    if ($allTrustedCerts) {
        foreach ($cert in $allTrustedCerts | Select-Object -First 3) {
            Write-Host "    - $($cert.Subject) [Thumbprint: $($cert.Thumbprint.Substring(0,16))...]" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "    (Nenhum certificado encontrado)" -ForegroundColor DarkGray
    }
}
Write-Host ""

# 8. Verificar se o plugin está instalado
Write-Host "[8/8] Verificando instalação do plugin..." -ForegroundColor Yellow
$pluginPath = "HKCU:\SOFTWARE\Microsoft\Office\Excel\Addins\VortexExcelAddIn"
if (Test-Path $pluginPath) {
    $loadBehavior = (Get-ItemProperty $pluginPath -ErrorAction SilentlyContinue).LoadBehavior
    Write-Host "  ✓ Plugin registrado" -ForegroundColor Green
    Write-Host "    LoadBehavior: $loadBehavior" -ForegroundColor Gray

    if ($loadBehavior -eq 3) {
        Write-Host "    Status: Carregado e Ativo" -ForegroundColor Green
    } elseif ($loadBehavior -eq 2) {
        Write-Host "    ⚠ Status: Desabilitado pelo usuário" -ForegroundColor Yellow
        $warnings++
    } elseif ($loadBehavior -eq 0) {
        Write-Host "    ⚠ Status: Desconectado" -ForegroundColor Yellow
        $warnings++
    } else {
        Write-Host "    ⚠ Status: Desconhecido ($loadBehavior)" -ForegroundColor Yellow
        $warnings++
    }
} else {
    Write-Host "  ⚠ Plugin não encontrado no registro" -ForegroundColor Yellow
    Write-Host "    O plugin pode não estar instalado ou pode estar instalado para todos os usuários" -ForegroundColor Gray
    $warnings++
}
Write-Host ""

# Verificar logs do plugin
Write-Host "Verificando logs do plugin..." -ForegroundColor Yellow
$logPath = "$env:APPDATA\VortexExcelAddIn\logs"
if (Test-Path $logPath) {
    $latestLog = Get-ChildItem $logPath -Filter "vortex-*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) {
        Write-Host "  Log mais recente: $($latestLog.Name)" -ForegroundColor Gray
        Write-Host "  Data: $($latestLog.LastWriteTime)" -ForegroundColor Gray
        Write-Host "  Tamanho: $([math]::Round($latestLog.Length/1KB, 2)) KB" -ForegroundColor Gray

        # Verificar últimas 10 linhas por erros
        $logContent = Get-Content $latestLog.FullName -Tail 20
        $errorLines = $logContent | Where-Object { $_ -match "ERROR|FATAL" }
        if ($errorLines) {
            Write-Host "  ⚠ Erros encontrados no log:" -ForegroundColor Yellow
            foreach ($line in $errorLines) {
                Write-Host "    $line" -ForegroundColor Red
            }
            $warnings++
        } else {
            Write-Host "  ✓ Nenhum erro recente encontrado no log" -ForegroundColor Green
        }
    } else {
        Write-Host "  ⚠ Nenhum arquivo de log encontrado" -ForegroundColor Yellow
        $warnings++
    }
} else {
    Write-Host "  ⚠ Pasta de logs não encontrada: $logPath" -ForegroundColor Yellow
    Write-Host "    Isso pode significar que o plugin nunca foi executado" -ForegroundColor Gray
    $warnings++
}
Write-Host ""

# Resumo
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "RESUMO DO DIAGNÓSTICO" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

if ($errors -eq 0 -and $warnings -eq 0) {
    Write-Host "✓ Tudo está OK! O plugin deve funcionar corretamente." -ForegroundColor Green
} elseif ($errors -eq 0) {
    Write-Host "⚠ $warnings aviso(s) encontrado(s)." -ForegroundColor Yellow
    Write-Host "  O plugin pode funcionar, mas verifique os avisos acima." -ForegroundColor Yellow
} else {
    Write-Host "✗ $errors erro(s) encontrado(s)." -ForegroundColor Red
    if ($warnings -gt 0) {
        Write-Host "⚠ $warnings aviso(s) encontrado(s)." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Corrija os erros acima antes de usar o plugin." -ForegroundColor Red
}

Write-Host ""
Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
