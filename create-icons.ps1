# Script para criar ícones de relógio para Auto-Refresh
Add-Type -AssemblyName System.Drawing

function Create-ClockIcon {
    param (
        [string]$outputPath,
        [System.Drawing.Color]$clockColor
    )

    # Criar bitmap 32x32
    $bitmap = New-Object System.Drawing.Bitmap(32, 32)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

    # Fundo transparente
    $graphics.Clear([System.Drawing.Color]::Transparent)

    # Desenhar círculo do relógio
    $pen = New-Object System.Drawing.Pen($clockColor, 2)
    $graphics.DrawEllipse($pen, 2, 2, 28, 28)

    # Desenhar ponteiro das horas (para 10:10)
    $graphics.DrawLine($pen, 16, 16, 16, 9)

    # Desenhar ponteiro dos minutos
    $graphics.DrawLine($pen, 16, 16, 23, 16)

    # Desenhar ponto central
    $brush = New-Object System.Drawing.SolidBrush($clockColor)
    $graphics.FillEllipse($brush, 14, 14, 4, 4)

    # Salvar
    $bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)

    # Cleanup
    $graphics.Dispose()
    $bitmap.Dispose()
    $pen.Dispose()
    $brush.Dispose()
}

# Criar ícone cinza (inativo)
$grayColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
Create-ClockIcon -outputPath "$PSScriptRoot\VortexExcelAddIn\Resources\RefreshIcon.png" -clockColor $grayColor

# Criar ícone verde (ativo)
$greenColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
Create-ClockIcon -outputPath "$PSScriptRoot\VortexExcelAddIn\Resources\RefreshActiveIcon.png" -clockColor $greenColor

Write-Host "Ícones criados com sucesso!" -ForegroundColor Green
Write-Host "- RefreshIcon.png (cinza - inativo)" -ForegroundColor Gray
Write-Host "- RefreshActiveIcon.png (verde - ativo)" -ForegroundColor Green
