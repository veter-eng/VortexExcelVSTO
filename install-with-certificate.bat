@echo off
echo ========================================
echo Vortex Excel Add-In - Instalacao Completa
echo (com Certificado de Seguranca)
echo ========================================
echo.

REM Verificar se está rodando como administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERRO] Este script precisa ser executado como Administrador!
    echo.
    echo Clique com botao direito no arquivo e selecione "Executar como administrador"
    pause
    exit /b 1
)

echo [OK] Executando como administrador
echo.

REM Fechar Excel
echo [1/8] Fechando Excel...
taskkill /F /IM EXCEL.EXE >NUL 2>&1
timeout /t 2 /nobreak >NUL
echo    [OK] Excel fechado
echo.

REM Limpar registros de itens desabilitados
echo [2/8] Limpando itens desabilitados...
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems" /f >NUL 2>&1
echo    [OK] Registros limpos
echo.

REM Limpar cache
echo [3/8] Limpando cache...
del /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*" >NUL 2>&1
echo    [OK] Cache limpo
echo.

REM Verificar se o certificado existe
echo [4/8] Verificando certificado...
cd /d "%~dp0"
if not exist "VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx" (
    echo    [ERRO] Certificado nao encontrado!
    echo    Arquivo esperado: VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx
    pause
    exit /b 1
)
echo    [OK] Certificado encontrado
echo.

REM Instalar certificado em Trusted Publishers
echo [5/8] Instalando certificado em Trusted Publishers...
powershell -Command "$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx'); $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('TrustedPublisher', 'CurrentUser'); $store.Open('ReadWrite'); $store.Add($cert); $store.Close()" >NUL 2>&1

if %errorLevel% equ 0 (
    echo    [OK] Certificado instalado com sucesso
) else (
    echo    [AVISO] Falha ao instalar certificado automaticamente
    echo    Tentando metodo alternativo...

    REM Método alternativo: exportar para .cer e importar
    powershell -Command "$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx'); [System.IO.File]::WriteAllBytes('temp_cert.cer', $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))" >NUL 2>&1
    certutil -addstore TrustedPublisher temp_cert.cer >NUL 2>&1
    del temp_cert.cer >NUL 2>&1

    if %errorLevel% equ 0 (
        echo    [OK] Certificado instalado via metodo alternativo
    ) else (
        echo    [ERRO] Nao foi possivel instalar o certificado
        echo    Voce precisara instalar manualmente via certmgr.msc
        pause
    )
)
echo.

REM Desinstalar versão antiga
echo [6/8] Desinstalando versao anterior...
if exist "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /uninstall "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" /silent
    timeout /t 2 /nobreak >NUL
    echo    [OK] Desinstalacao concluida
) else (
    echo    [AVISO] Arquivo .vsto nao encontrado
)
echo.

REM Instalar
echo [7/8] Instalando plugin...
if exist "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    echo    Abrindo instalador...
    start "" "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto"
    echo    [OK] Instalador aberto!
    echo.
    echo    IMPORTANTE: Clique em "Instalar" na janela que abriu
    echo    O certificado ja foi instalado, entao NAO deve aparecer aviso de seguranca
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
echo [8/8] Abrindo Excel para teste...
echo ========================================
echo.
echo Testando instalacao...
echo 1. Verifique se a aba "Vortex" aparece no Ribbon
echo 2. Clique no botao "Vortex Plugin"
echo 3. A barra lateral "Vortex Data Plugin" deve aparecer a direita
echo.

start excel

echo.
echo ========================================
echo Instalacao Concluida!
echo ========================================
echo.
echo Se a barra lateral NAO abrir:
echo 1. Execute o script diagnose-plugin.ps1 para verificar o problema
echo 2. Verifique os logs em: %%APPDATA%%\VortexExcelAddIn\logs
echo.
pause
