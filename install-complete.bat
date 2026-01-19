@echo off
echo ========================================
echo Vortex Excel Add-In - Instalacao COMPLETA
echo Resolve TODOS os problemas do diagnostico
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

cd /d "%~dp0"

echo ========================================
echo ETAPA 1: Verificando Pre-requisitos
echo ========================================
echo.

REM Verificar se VSTO Runtime está instalado
echo [1.1] Verificando VSTO Runtime...
if exist "%ProgramFiles%\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" (
    echo    [OK] VSTO Runtime ja instalado
) else (
    echo    [AVISO] VSTO Runtime NAO encontrado!
    echo    Voce precisa instalar o VSTO Runtime antes de continuar.
    echo.
    echo    Passos:
    echo    1. Baixe o instalador em:
    echo       https://www.microsoft.com/download/details.aspx?id=56961
    echo    2. Execute o instalador baixado
    echo    3. REINICIE o computador
    echo    4. Execute este script novamente
    echo.

    choice /C SN /M "Abrir pagina de download agora"
    if errorlevel 2 goto skip_vsto
    if errorlevel 1 start https://www.microsoft.com/download/details.aspx?id=56961

    :skip_vsto
    echo.
    echo Instale o VSTO Runtime e execute este script novamente.
    pause
    exit /b 1
)
echo.

REM Verificar se o certificado existe
echo [1.2] Verificando certificado...
if not exist "VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx" (
    echo    [ERRO] Certificado nao encontrado!
    echo    Arquivo esperado: VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx
    echo.
    echo    Certifique-se de que voce esta executando este script na pasta raiz do projeto.
    pause
    exit /b 1
)
echo    [OK] Certificado encontrado
echo.

REM Verificar se o .vsto existe
echo [1.3] Verificando arquivo de instalacao...
if not exist "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" (
    echo    [ERRO] Arquivo de instalacao nao encontrado!
    echo    Voce precisa compilar o projeto primeiro.
    echo.
    echo    Execute: msbuild VortexExcelAddIn\VortexExcelAddIn.csproj /p:Configuration=Release
    pause
    exit /b 1
)
echo    [OK] Arquivo de instalacao encontrado
echo.

echo ========================================
echo ETAPA 2: Preparando Ambiente
echo ========================================
echo.

REM Fechar Excel
echo [2.1] Fechando Excel...
taskkill /F /IM EXCEL.EXE >NUL 2>&1
timeout /t 2 /nobreak >NUL
echo    [OK] Excel fechado
echo.

REM Limpar registros de itens desabilitados
echo [2.2] Limpando itens desabilitados do Excel...
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems" /f >NUL 2>&1
reg delete "HKCU\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems" /f >NUL 2>&1
echo    [OK] Registros limpos
echo.

REM Limpar cache
echo [2.3] Limpando cache do Office...
del /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*" >NUL 2>&1
del /Q "%LOCALAPPDATA%\Microsoft\Office\15.0\Wef\*.*" >NUL 2>&1
echo    [OK] Cache limpo
echo.

echo ========================================
echo ETAPA 3: Instalando Certificado
echo ========================================
echo.

echo [3.1] Instalando certificado de seguranca em Trusted Publishers...
powershell -Command "$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx'); $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('TrustedPublisher', 'CurrentUser'); $store.Open('ReadWrite'); $store.Add($cert); $store.Close()" >NUL 2>&1

if %errorLevel% equ 0 (
    echo    [OK] Certificado instalado com sucesso
) else (
    echo    [AVISO] Falha ao instalar certificado via PowerShell
    echo    Tentando metodo alternativo...

    REM Método alternativo: exportar para .cer e importar via certutil
    powershell -Command "$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx'); [System.IO.File]::WriteAllBytes('temp_cert.cer', $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))" >NUL 2>&1
    certutil -addstore TrustedPublisher temp_cert.cer >NUL 2>&1
    del temp_cert.cer >NUL 2>&1

    if %errorLevel% equ 0 (
        echo    [OK] Certificado instalado via certutil
    ) else (
        echo    [ERRO] Nao foi possivel instalar o certificado automaticamente
        echo.
        echo    INSTALACAO MANUAL NECESSARIA:
        echo    1. Pressione Win + R e digite: certmgr.msc
        echo    2. Expanda "Editores Confiaveis"
        echo    3. Clique com botao direito em "Certificados" -^> "Todas as Tarefas" -^> "Importar..."
        echo    4. Navegue ate VortexExcelAddIn\VortexExcelAddIn_TemporaryKey.pfx
        echo    5. Complete o assistente
        echo.
        pause
        echo Continuando...
    )
)
echo.

echo ========================================
echo ETAPA 4: Instalando Plugin
echo ========================================
echo.

echo [4.1] Desinstalando versao anterior (se existir)...
"C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /uninstall "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto" /silent >NUL 2>&1
timeout /t 2 /nobreak >NUL
echo    [OK] Desinstalacao concluida
echo.

echo [4.2] Instalando nova versao...
echo    Abrindo instalador do plugin...
start "" "VortexExcelAddIn\bin\Release\VortexExcelAddIn.vsto"
echo    [OK] Instalador aberto!
echo.
echo    *** IMPORTANTE ***
echo    Na janela que abriu:
echo    1. Clique em "Instalar"
echo    2. Com o certificado instalado, NAO deve aparecer aviso de seguranca
echo    3. Aguarde a mensagem "Instalacao bem-sucedida"
echo.

echo Aguarde 10 segundos para a instalacao concluir...
timeout /t 10 /nobreak

echo.
echo ========================================
echo ETAPA 5: Testando Instalacao
echo ========================================
echo.

echo [5.1] Abrindo Excel para teste...
start excel
echo    [OK] Excel aberto
echo.

echo ========================================
echo INSTALACAO CONCLUIDA!
echo ========================================
echo.
echo Agora teste o plugin:
echo.
echo 1. No Excel, procure a aba "Vortex" no Ribbon
echo 2. Clique no botao "Vortex Plugin"
echo 3. A barra lateral "Vortex Data Plugin" deve aparecer a direita
echo.
echo Se a barra lateral NAO aparecer:
echo 1. Feche o Excel completamente
echo 2. Execute o script diagnose-plugin.ps1 novamente
echo 3. Envie o resultado do diagnostico para suporte
echo.
echo Logs do plugin em: %%APPDATA%%\VortexExcelAddIn\logs
echo.
pause
