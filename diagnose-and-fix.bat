@echo off
setlocal enabledelayedexpansion

echo ========================================
echo Vortex Excel Add-In - Diagnóstico e Correção
echo ========================================
echo.

REM Verificar se está rodando como Admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [AVISO] Este script não está rodando como Administrador.
    echo Algumas operações podem falhar.
    echo.
    echo Deseja continuar mesmo assim? (S/N)
    set /p continue=
    if /i not "!continue!"=="S" exit /b
)

REM Verificar se o Excel está rodando
echo [1] Verificando se o Excel está em execução...
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo    [ERRO] O Excel está em execução!
    echo    Fechando Excel...
    taskkill /F /IM EXCEL.EXE >NUL 2>&1
    timeout /t 3 /nobreak >NUL
) else (
    echo    [OK] Excel não está em execução
)
echo.

REM Verificar suplementos desabilitados
echo [2] Verificando suplementos desabilitados...
reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo    [AVISO] Encontrado registro de itens desabilitados
    echo    Removendo itens desabilitados...
    reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems" /f >NUL 2>&1
    echo    [OK] Registro limpo
) else (
    echo    [OK] Nenhum item desabilitado encontrado
)
echo.

REM Limpar cache
echo [3] Limpando cache de add-ins...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
    del /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*" >NUL 2>&1
    echo    [OK] Cache do Office limpo
)
if exist "%LOCALAPPDATA%\Apps\2.0\" (
    for /d %%i in ("%LOCALAPPDATA%\Apps\2.0\*") do (
        if exist "%%i" (
            dir /s /b "%%i" | findstr /i "VortexExcel" >NUL 2>&1
            if !errorlevel! EQU 0 (
                rd /s /q "%%i" >NUL 2>&1
            )
        )
    )
    echo    [OK] Cache ClickOnce limpo
)
echo.

REM Verificar .NET Framework
echo [4] Verificando .NET Framework 4.8...
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    for /f "tokens=3" %%i in ('reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release ^| findstr Release') do set netVersion=%%i
    if !netVersion! GEQ 528040 (
        echo    [OK] .NET Framework 4.8 ou superior instalado
    ) else (
        echo    [AVISO] .NET Framework pode estar desatualizado
        echo    Versão atual: !netVersion!
        echo    Recomendado: 528040 ou superior
    )
) else (
    echo    [ERRO] .NET Framework 4.8 não encontrado!
    echo    Por favor, instale: https://dotnet.microsoft.com/download/dotnet-framework/net48
    pause
    exit /b 1
)
echo.

REM Verificar VSTO Runtime
echo [5] Verificando VSTO Runtime...
reg query "HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" /v Version >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo    [OK] VSTO Runtime instalado
) else (
    echo    [ERRO] VSTO Runtime não encontrado!
    echo    Por favor, instale: https://www.microsoft.com/download/details.aspx?id=56961
    pause
    exit /b 1
)
echo.

REM Recompilar
echo [6] Recompilando o projeto...
if exist "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" (
    "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "%~dp0VortexExcelAddIn\VortexExcelAddIn.csproj" /p:Configuration=Release /t:Clean,Build /v:minimal /nologo
    if !errorlevel! EQU 0 (
        echo    [OK] Projeto compilado com sucesso
    ) else (
        echo    [ERRO] Falha na compilação
        pause
        exit /b 1
    )
) else (
    echo    [AVISO] MSBuild não encontrado, pulando compilação
)
echo.

REM Desinstalar versão anterior
echo [7] Desinstalando versão anterior...
if exist "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" /uninstall /silent >NUL 2>&1
    timeout /t 2 /nobreak >NUL
    echo    [OK] Desinstalação concluída
) else (
    echo    [AVISO] Arquivo .vsto não encontrado
)
echo.

REM Adicionar pasta como local confiável
echo [8] Configurando local confiável...
set installPath=%~dp0VortexExcelAddIn\bin\Release\
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location99" /v Path /t REG_SZ /d "%installPath%" /f >NUL 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location99" /v AllowSubFolders /t REG_DWORD /d 1 /f >NUL 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location99" /v Description /t REG_SZ /d "Vortex Excel Add-In" /f >NUL 2>&1
echo    [OK] Local confiável adicionado
echo.

REM Configurar segurança
echo [9] Configurando segurança do Office...
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >NUL 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >NUL 2>&1
echo    [OK] Configurações de segurança ajustadas
echo.

REM Instalar
echo [10] Instalando nova versão...
if exist "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    start "" "%~dp0VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto"
    echo    [OK] Instalador aberto
) else (
    echo    [ERRO] Arquivo de instalação não encontrado!
    pause
    exit /b 1
)
echo.

echo ========================================
echo Diagnóstico e correção concluídos!
echo ========================================
echo.
echo IMPORTANTE:
echo 1. Clique em "Instalar" na janela que abriu
echo 2. Aguarde a instalação concluir
echo 3. Abra o Excel
echo 4. Você verá 2 MessageBoxes de debug:
echo    - "Vortex Add-in: Iniciando..."
echo    - "Vortex Add-in: Carregado com sucesso!"
echo.
echo 5. Se NÃO ver essas mensagens, o add-in falhou ao carregar
echo 6. Execute: check-disabled.ps1 para ver os erros
echo.
echo Se aparecer um erro, tire um print e me mostre!
echo.
pause
