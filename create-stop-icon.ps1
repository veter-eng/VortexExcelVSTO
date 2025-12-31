# Script para criar ícone de Stop para o menu Auto-Refresh
Add-Type -AssemblyName System.Drawing

# Criar ícone de stop (16x16 para menu item)
$stopIcon = New-Object System.Drawing.Bitmap(16, 16)
$graphics = [System.Drawing.Graphics]::FromImage($stopIcon)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

# Fundo transparente
$graphics.Clear([System.Drawing.Color]::Transparent)

# Quadrado vermelho no centro (símbolo de stop)
$stopBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(231, 76, 60)) # #E74C3C
$stopRect = New-Object System.Drawing.Rectangle(3, 3, 10, 10)
$graphics.FillRectangle($stopBrush, $stopRect)

# Borda mais escura
$borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(192, 57, 43), 1) # #C0392B
$graphics.DrawRectangle($borderPen, $stopRect)

# Salvar
$stopIconPath = "C:\Users\rikea\RiderProjects\VortexExcelVSTO\VortexExcelAddIn\Resources\StopIcon.png"
$stopIcon.Save($stopIconPath, [System.Drawing.Imaging.ImageFormat]::Png)

Write-Host "Ícone de stop criado: $stopIconPath" -ForegroundColor Green

# Cleanup
$graphics.Dispose()
$stopIcon.Dispose()
$stopBrush.Dispose()
$borderPen.Dispose()
