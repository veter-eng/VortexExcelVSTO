# Script para verificar add-ins desabilitados no Excel
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Verificando Add-ins Desabilitados" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Verificar no Registry
$paths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\VortexExcelAddIn",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\14.0\Excel\Resiliency\DisabledItems"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Write-Host "[ENCONTRADO] $path" -ForegroundColor Yellow
        Get-ItemProperty -Path $path | Format-List
    }
}

# Verificar Event Viewer para erros do VSTO
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Verificando Logs de Erro (últimas 24h)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$errors = Get-EventLog -LogName Application -After (Get-Date).AddDays(-1) -EntryType Error |
    Where-Object { $_.Source -like "*VSTO*" -or $_.Message -like "*VortexExcel*" -or $_.Message -like "*Office*" } |
    Select-Object -First 10

if ($errors) {
    foreach ($error in $errors) {
        Write-Host "[ERRO]" -ForegroundColor Red
        Write-Host "  Fonte: $($error.Source)" -ForegroundColor Yellow
        Write-Host "  Hora: $($error.TimeGenerated)" -ForegroundColor Yellow
        Write-Host "  Mensagem: $($error.Message)" -ForegroundColor White
        Write-Host ""
    }
} else {
    Write-Host "[OK] Nenhum erro relacionado ao VSTO encontrado nos logs." -ForegroundColor Green
}

Write-Host "`nPressione qualquer tecla para continuar..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
