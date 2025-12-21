# Solução de Problemas - Vortex Excel Add-In

## Problema: Plugin aparece no COM Add-ins mas não carrega

### Sintomas:
- O plugin aparece na lista de Suplementos COM
- Quando marcado e clicado em OK, nada acontece
- O campo "Location" mostra `file:///C:/...`
- Não aparece nenhum botão no Ribbon

### Solução 1: Adicionar Local de Rede Confiável

1. **Abra o Excel**
2. Vá em **Arquivo → Opções**
3. Selecione **Central de Confiabilidade**
4. Clique em **Configurações da Central de Confiabilidade**
5. Selecione **Locais Confiáveis**
6. Clique em **Adicionar novo local...**
7. Digite o caminho: `C:\Users\[SEU_USUÁRIO]\RiderProjects\VortexExcelVSTO\VortexExcelAddIn\bin\Release\`
8. ✅ Marque **"As subpastas deste local também são confiáveis"**
9. Clique em **OK** em todas as janelas
10. Feche e reabra o Excel

### Solução 2: Ajustar Configurações de Macro

1. **Abra o Excel**
2. Vá em **Arquivo → Opções**
3. Selecione **Central de Confiabilidade**
4. Clique em **Configurações da Central de Confiabilidade**
5. Selecione **Configurações de Macro**
6. Selecione **"Habilitar todas as macros"** (apenas para teste)
7. ✅ Marque **"Confiar no acesso ao modelo de objeto do projeto VBA"**
8. Clique em **OK**
9. Feche e reabra o Excel

### Solução 3: Desabilitar Proteção de Suplementos (Apenas Desenvolvimento)

1. **Abra o Excel**
2. Vá em **Arquivo → Opções**
3. Selecione **Central de Confiabilidade**
4. Clique em **Configurações da Central de Confiabilidade**
5. Selecione **Configurações de Suplemento**
6. ❌ **Desmarque** "Exigir que as Extensões de Aplicativo sejam assinadas por um Fornecedor Confiável"
7. ❌ **Desmarque** "Desabilitar todas as Extensões de Aplicativo"
8. Clique em **OK**
9. Feche e reabra o Excel

### Solução 4: Limpar Cache de Add-ins

Às vezes o Excel cacheia informações antigas do add-in.

**Windows:**
```cmd
# Feche o Excel primeiro
del /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*"
del /Q "%LOCALAPPDATA%\Apps\2.0\*VortexExcel*" /S
```

**PowerShell:**
```powershell
# Feche o Excel primeiro
Remove-Item "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Force -ErrorAction SilentlyContinue
Get-ChildItem "$env:LOCALAPPDATA\Apps\2.0" -Recurse -Filter "*VortexExcel*" | Remove-Item -Force -Recurse
```

Depois:
1. Desinstale o add-in pelo Painel de Controle
2. Reinstale executando `VortexExcelAddIn.vsto`
3. Abra o Excel

### Solução 5: Verificar .NET Framework

O add-in requer .NET Framework 4.8 ou superior.

**Verificar versão instalada:**
1. Abra o Prompt de Comando
2. Execute:
   ```cmd
   reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release
   ```
3. Se o valor for menor que 528040, você precisa atualizar

**Baixar .NET Framework 4.8:**
https://dotnet.microsoft.com/download/dotnet-framework/net48

### Solução 6: Reinstalar VSTO Runtime

Se nada funcionar, reinstale o VSTO Runtime:

1. **Desinstale** o VSTO Runtime atual:
   - Painel de Controle → Programas → Desinstalar
   - Procure por "Microsoft Visual Studio 2010 Tools for Office Runtime"
   - Desinstale

2. **Baixe e instale** a versão mais recente:
   - https://www.microsoft.com/en-us/download/details.aspx?id=56961
   - Execute o instalador
   - Reinicie o computador

3. **Reinstale o add-in**

### Solução 7: Verificar Logs de Erro

O add-in gera logs que podem ajudar a identificar o problema:

1. Pressione `Win + R`
2. Digite: `%AppData%\VortexExcelAddIn\logs`
3. Abra o arquivo de log mais recente
4. Procure por linhas com "ERROR" ou "FATAL"

Se não houver pasta de logs, o add-in não está sendo carregado.

### Solução 8: Forçar Registro Manual (Avançado)

Se você é desenvolvedor e nada mais funcionar:

1. Abra o **Prompt de Comando como Administrador**
2. Execute:
   ```cmd
   cd "C:\Users\[SEU_USUÁRIO]\RiderProjects\VortexExcelVSTO\VortexExcelAddIn\bin\Release"

   "C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /Install VortexExcelAddIn.vsto
   ```

### Solução 9: Verificar se o Ribbon está carregado

Se o add-in carrega mas o botão não aparece:

1. Abra o Excel
2. Clique com o botão direito no Ribbon
3. Selecione **"Personalizar o Friso..."** ou **"Customize the Ribbon..."**
4. No lado direito, verifique se há uma aba **"Suplementos"** ou **"Add-Ins"**
5. Expanda e veja se **"Vortex Data"** está listado
6. Se estiver desmarcado, marque-o

## Como saber se o add-in está funcionando corretamente

### Checklist:

✅ O add-in aparece em **Arquivo → Opções → Suplementos → Suplementos COM**
✅ O add-in está **marcado** na lista
✅ Não há mensagem de erro ao marcar
✅ Após reabrir o Excel, aparece uma aba **"Suplementos"** ou **"Add-Ins"** no Ribbon
✅ Dentro dessa aba, há um grupo **"Vortex Data"**
✅ Dentro do grupo, há um botão **"Vortex Plugin"**
✅ Ao clicar no botão, o painel lateral aparece/desaparece

## Testando o Plugin

Depois de instalar corretamente:

1. **Abra o Excel**
2. Clique na aba **"Suplementos"** ou **"Add-Ins"** no Ribbon
3. Clique no botão **"Vortex Plugin"** no grupo "Vortex Data"
4. Um painel lateral deve aparecer à direita
5. Configure a conexão com InfluxDB na aba "Configuração"
6. Teste a conexão
7. Vá para a aba "Consulta" e faça uma consulta

## Ambiente de Desenvolvimento vs Produção

### Desenvolvimento (Visual Studio):
- Use **F5** para depurar
- O add-in é instalado temporariamente
- Certificado temporário é usado
- Logs mais detalhados

### Produção (Instalação Manual):
- Use o arquivo `.vsto` para instalar
- Para distribuir, você deve:
  1. Assinar com um certificado real (não temporário)
  2. Publicar em um servidor web ou compartilhamento de rede
  3. Criar um instalador (ClickOnce ou MSI)

## Contato para Suporte

Se nenhuma solução funcionar:
1. Capture uma captura de tela da lista de Suplementos COM
2. Copie o conteúdo do arquivo de log (se existir)
3. Anote a versão do Excel (Arquivo → Conta → Sobre o Excel)
4. Abra uma issue com essas informações
