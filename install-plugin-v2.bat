@echo off
echo ========================================
echo Vortex Excel Add-In - Instalacao v2
echo ========================================
echo.

REM Fechar Excel
echo [1] Fechando Excel...
taskkill /F /IM EXCEL.EXE >NUL 2>&1
timeout /t 2 /nobreak >NUL
echo    [OK] Excel fechado
echo.

REM Limpar registros de itens desabilitados
echo [2] Limpando itens desabilitados...
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems" /f >NUL 2>&1
echo    [OK] Registros limpos
echo.

REM Limpar cache
echo [3] Limpando cache...
del /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*" >NUL 2>&1
echo    [OK] Cache limpo
echo.

REM Pular configuracao de local confiavel (pode causar travamento)
echo [4] Pulando configuracao de local confiavel...
echo    [OK] Pulado
echo.

REM Desinstalar versao antiga
echo [5] Desinstalando versao anterior...
cd /d "%~dp0"
if exist "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /uninstall "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" /silent
    timeout /t 2 /nobreak >NUL
    echo    [OK] Desinstalacao concluida
) else (
    echo    [AVISO] Arquivo .vsto nao encontrado
)
echo.

REM Instalar
echo [6] Instalando...
if exist "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    echo    Abrindo instalador...
    start "" "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto"
    echo    [OK] Instalador aberto!
    echo.
    echo    Clique em "Instalar" na janela que abriu
    echo    Aguarde a instalacao concluir...
) else (
    echo    [ERRO] Arquivo nao encontrado!
    pause
    exit /b 1
)
echo.

echo Aguarde 5 segundos para a instalacao concluir...
timeout /t 5 /nobreak

echo.
echo ========================================
echo Abrindo Excel...
echo ========================================
echo.
echo IMPORTANTE: O plugin agora inicia silenciosamente.
echo Verifique se a aba "Suplementos" aparece no Ribbon.
echo.

start excel

echo.
echo Se NAO aparecerem as mensagens:
echo 1. Va em Arquivo -^> Opcoes -^> Suplementos
echo 2. Verifique "Suplementos Desabilitados"
echo 3. Habilite o VortexExcelAddIn se estiver la
echo.
pause
