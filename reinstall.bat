@echo off
echo ========================================
echo Vortex Excel Add-In - Reinstalação
echo ========================================
echo.

REM Verificar se o Excel está rodando
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo [ERRO] O Excel está em execução!
    echo Por favor, feche o Excel e execute este script novamente.
    echo.
    pause
    exit /b 1
)

echo [1/4] Limpando cache de add-ins...
del /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*" 2>NUL
del /Q "%LOCALAPPDATA%\Apps\2.0\*VortexExcel*" /S 2>NUL
echo       Cache limpo!
echo.

echo [2/4] Compilando o projeto...
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "%~dp0VortexExcelAddIn\VortexExcelAddIn.csproj" /p:Configuration=Release /t:Clean,Build /v:minimal
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Falha na compilação!
    pause
    exit /b 1
)
echo       Compilado com sucesso!
echo.

echo [3/4] Desinstalando versão anterior (se existir)...
"%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" /uninstall /silent 2>NUL
timeout /t 2 /nobreak >NUL
echo       Desinstalação concluída!
echo.

echo [4/4] Instalando nova versão...
"%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto"
echo.

echo ========================================
echo Instalação concluída!
echo ========================================
echo.
echo Próximos passos:
echo 1. Abra o Microsoft Excel
echo 2. Procure pela aba "Suplementos" ou "Add-Ins"
echo 3. Clique no botão "Vortex Plugin"
echo 4. O painel lateral deve aparecer!
echo.
echo Se tiver problemas, consulte: TROUBLESHOOTING.md
echo.
pause
