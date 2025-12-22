@echo off
echo ========================================
echo Vortex Excel Add-In - Instalacao Forcada
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
reg delete "HKCU\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems" /f >NUL 2>&1
echo    [OK] Registros limpos
echo.

REM Limpar cache
echo [3] Limpando cache...
del /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*" >NUL 2>&1
del /Q "%LOCALAPPDATA%\Microsoft\Office\15.0\Wef\*.*" >NUL 2>&1
echo    [OK] Cache limpo
echo.

REM Adicionar local confiavel
echo [4] Adicionando local confiavel...
set installPath=%~dp0VortexExcelAddIn\bin\Release\
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location99" /v Path /t REG_SZ /d "%installPath%" /f >NUL 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location99" /v AllowSubFolders /t REG_DWORD /d 1 /f >NUL 2>&1
echo    [OK] Local confiavel adicionado: %installPath%
echo.

REM Configurar seguranca
echo [5] Configurando seguranca...
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >NUL 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >NUL 2>&1
echo    [OK] Seguranca configurada
echo.

REM Desinstalar versao antiga
echo [6] Desinstalando versao anterior (se existir)...
if exist "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /uninstall "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" /silent >NUL 2>&1
    timeout /t 2 /nobreak >NUL
    echo    [OK] Desinstalacao concluida
) else (
    echo    [AVISO] Arquivo .vsto nao encontrado
)
echo.

REM Instalar via VSTOInstaller
echo [7] Instalando via VSTOInstaller.exe...
if exist "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    echo    Instalando de: %~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /install "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" /silent

    if %errorlevel% EQU 0 (
        echo    [OK] Instalacao concluida!
    ) else (
        echo    [ERRO] Codigo de erro: %errorlevel%
        echo.
        echo    Tentando instalacao interativa...
        start "" "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto"
    )
) else (
    echo    [ERRO] Arquivo nao encontrado!
    pause
    exit /b 1
)
echo.

echo ========================================
echo Instalacao concluida!
echo ========================================
echo.
echo IMPORTANTE:
echo 1. Abra o Excel agora
echo 2. Voce DEVE ver 2 MessageBoxes:
echo    - "Vortex Add-in: Iniciando..."
echo    - "Vortex Add-in: Carregado com sucesso!"
echo.
echo 3. Se as mensagens NAO aparecerem:
echo    - Va em Arquivo -^> Opcoes -^> Suplementos
echo    - No final, em "Gerenciar:", selecione "Suplementos Desabilitados"
echo    - Clique em "Ir..." e veja se VortexExcelAddIn esta la
echo.
echo 4. Se estiver em Desabilitados, clique em "Habilitar"
echo.
pause

echo.
echo Abrindo Excel agora...
start excel
