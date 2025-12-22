# Vortex Excel Add-In

Plugin VSTO para Microsoft Excel que permite consultar e importar dados do InfluxDB diretamente no Excel.

## Características

- 📊 Consulta de dados do InfluxDB com interface intuitiva
- 🔄 Importação automática de dados para planilhas Excel
- 🎯 Filtros em cascata (Coletor → Gateway → Equipamento → Tag)
- 📅 Seleção de período com data/hora de início e fim
- 💾 Exportação de dados para CSV
- 🎨 Interface WPF moderna integrada ao Excel

## Pré-requisitos

Antes de instalar o plugin, certifique-se de ter:

- ✅ **Microsoft Excel** (2013 ou superior)
- ✅ **Windows** 7 ou superior
- ✅ **.NET Framework 4.8** ou superior
- ✅ **Visual Studio 2010 Tools for Office Runtime** (VSTO Runtime)

### Instalando o VSTO Runtime

Se você não tem o VSTO Runtime instalado:

1. Baixe o instalador: [Microsoft Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=56961)
2. Execute o instalador baixado
3. Siga as instruções na tela

## Instalação do Plugin

### Método 1: Instalação Automática via Script (Recomendado)

1. **Execute o script de instalação:**
   ```bash
   install-plugin-v2.bat
   ```

2. **O script irá automaticamente:**
   - ✅ Fechar o Excel se estiver aberto
   - ✅ Limpar itens desabilitados no registro
   - ✅ Limpar cache de add-ins
   - ✅ Desinstalar versões anteriores
   - ✅ Instalar a nova versão
   - ✅ Abrir o Excel

3. **Clique em "Instalar"** na janela que aparecer

4. **Verifique a instalação:**
   - Você deve ver **2 MessageBoxes** ao abrir o Excel:
     - "Vortex Add-in: Iniciando..."
     - "Vortex Add-in: Carregado com sucesso!"
   - No Ribbon do Excel, aparecerá uma aba chamada **"Vortex"**
   - Dentro da aba Vortex, haverá um botão **"Vortex Plugin"**

### Método 2: Instalação Manual via arquivo .vsto

1. **Compile o projeto:**
   ```bash
   msbuild VortexExcelAddIn\VortexExcelAddIn.csproj /p:Configuration=Release
   ```

2. **Localize o arquivo de instalação:**
   - Navegue até a pasta: `VortexExcelAddIn\bin\Release\`
   - Encontre o arquivo `VortexExcelAddIn.vsto`

3. **Execute o instalador:**
   - Clique duas vezes em `VortexExcelAddIn.vsto`
   - Uma janela de instalação será exibida

4. **Aceite o aviso de segurança:**
   - Clique em **"Instalar"** na janela de instalação
   - O plugin será instalado automaticamente

5. **Abra o Microsoft Excel:**
   - Você deve ver as 2 MessageBoxes de confirmação
   - Uma aba chamada **"Vortex"** aparecerá no Ribbon

### Método 3: Instalação via Visual Studio (Para desenvolvedores)

1. **Abra o projeto no Visual Studio:**
   ```bash
   start VortexExcelAddIn\VortexExcelAddIn.csproj
   ```

2. **Execute o projeto:**
   - Pressione **F5** ou clique em "Iniciar Depuração"
   - O Visual Studio irá compilar, instalar temporariamente o add-in e abrir o Excel

3. **Para instalação permanente:**
   - Compile em modo Release: **Build → Build Solution**
   - Siga as instruções do Método 1

## Verificando a Instalação

Após a instalação, verifique se o plugin está ativo:

1. Abra o **Microsoft Excel**
2. Vá em **Arquivo → Opções**
3. Selecione **Suplementos** no menu lateral
4. Na parte inferior da janela:
   - Em "Gerenciar:", selecione **"Suplementos COM"**
   - Clique em **"Ir..."**
5. Você deve ver **"VortexExcelAddIn"** na lista com uma ✅ marcação

## Usando o Plugin

### Primeira Configuração

1. **Abra o painel do plugin:**
   - No Excel, clique na aba **"Vortex"** no Ribbon
   - Clique no botão **"Vortex Plugin"**
   - O painel lateral "Vortex Data Plugin" será exibido à direita

2. **Configure a conexão com InfluxDB:**
   - Clique na aba **"Configuração"**
   - Preencha os dados de conexão:
     - **URL do InfluxDB**: `http://seu-servidor:8086`
     - **Token de Acesso**: Seu token de autenticação
     - **Organização**: Nome da sua organização
     - **Bucket**: Nome do bucket padrão
   - Clique em **"Testar Conexão"** para validar
   - Clique em **"Salvar"**

### Consultando Dados

1. **Acesse a aba "Consulta"**

2. **Selecione os filtros:**
   - **Coletor**: Escolha o coletor de dados
   - **Gateway**: Selecione o gateway (carregado automaticamente)
   - **Equipamento**: Escolha o equipamento (carregado automaticamente)
   - **Tag**: Selecione a tag desejada (carregada automaticamente)

3. **Defina o período:**
   - **Data/Hora Início**: Data e hora inicial da consulta
   - **Data/Hora Fim**: Data e hora final da consulta
   - **Limite de Registros**: Número máximo de resultados (padrão: 1000)

4. **Execute a consulta:**
   - Clique em **"Consultar"**
   - Os dados serão exibidos na visualização prévia

5. **Importe para o Excel:**
   - Clique em **"Inserir no Excel"**
   - Os dados serão inseridos na planilha ativa

### Exportando para CSV

1. Após realizar uma consulta com sucesso
2. Clique em **"Exportar CSV"**
3. Escolha o local para salvar o arquivo
4. O arquivo CSV será gerado com todos os dados da consulta

## Configuração do NLog (Logs)

O plugin gera logs de execução. Para configurar:

1. Crie um arquivo `NLog.config` na mesma pasta do Excel ou na pasta do usuário
2. Exemplo de configuração básica:

```xml
<?xml version="1.0" encoding="utf-8" ?>
<nlog xmlns="http://www.nlog-project.org/schemas/NLog.xsd"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <targets>
    <target name="file" xsi:type="File"
            fileName="${specialfolder:folder=ApplicationData}/VortexExcelAddIn/logs/vortex-${shortdate}.log"
            layout="${longdate} ${level:uppercase=true} ${message} ${exception:format=tostring}" />
  </targets>
  <rules>
    <logger name="*" minlevel="Info" writeTo="file" />
  </rules>
</nlog>
```

Os logs ficarão em: `%AppData%/VortexExcelAddIn/logs/`

## Desinstalação

### Opção 1: Painel de Controle

1. Abra **Painel de Controle**
2. Vá em **Programas → Programas e Recursos**
3. Procure por **"VortexExcelAddIn"** na lista
4. Clique com o botão direito e selecione **"Desinstalar"**
5. Siga as instruções na tela

### Opção 2: Via Excel

1. Abra o **Microsoft Excel**
2. Vá em **Arquivo → Opções → Suplementos**
3. Em "Gerenciar:", selecione **"Suplementos COM"** e clique em **"Ir..."**
4. Desmarque **"VortexExcelAddIn"**
5. Clique em **"OK"**

Nota: Isso apenas desabilita o plugin, não o remove completamente.

## Solução de Problemas

### As MessageBoxes de inicialização não aparecem

Se você não vê as mensagens "Vortex Add-in: Iniciando..." e "Vortex Add-in: Carregado com sucesso!":

**Solução 1: Verificar suplementos desabilitados**
1. Vá em **Arquivo → Opções → Suplementos**
2. No dropdown inferior, selecione **"Itens Desabilitados"** e clique em **"Ir..."**
3. Se "VortexExcelAddIn" estiver na lista, selecione-o e clique em **"Habilitar"**
4. Reinicie o Excel

**Solução 2: Usar o script de diagnóstico**
1. Execute o script `diagnose-and-fix.bat`
2. O script irá:
   - Verificar e limpar itens desabilitados
   - Limpar cache de add-ins
   - Recompilar e reinstalar o plugin
3. Siga as instruções na tela

### A aba "Vortex" não aparece no Ribbon

**Solução 1: Habilitar o plugin**
1. Vá em **Arquivo → Opções → Suplementos**
2. Verifique se "VortexExcelAddIn" está na lista
3. Se estiver desmarcado, marque-o
4. Se estiver em "Suplementos Desabilitados", mova-o para "Suplementos Ativos"

**Solução 2: Verificar a segurança**
1. Vá em **Arquivo → Opções → Central de Confiabilidade**
2. Clique em **"Configurações da Central de Confiabilidade"**
3. Selecione **"Configurações de Suplemento"**
4. Desmarque **"Exigir que as Extensões de Aplicativo sejam assinadas por um Fornecedor Confiável"** (apenas para desenvolvimento/teste)

### Erro ao conectar com InfluxDB

**Verifique:**
- ✅ A URL está correta (incluindo http:// ou https://)
- ✅ O token de acesso é válido
- ✅ O firewall não está bloqueando a conexão
- ✅ O InfluxDB está rodando e acessível

**Verifique os logs:**
- Vá em `%AppData%/VortexExcelAddIn/logs/`
- Abra o arquivo de log mais recente
- Procure por mensagens de erro

### Erro "VSTO Runtime não encontrado"

1. Instale o [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=56961)
2. Reinicie o computador
3. Tente instalar o plugin novamente

### O painel do plugin não abre

1. Feche o Excel completamente
2. Abra o Gerenciador de Tarefas (Ctrl+Shift+Esc)
3. Certifique-se de que não há processos do Excel em execução
4. Abra o Excel novamente

### Erro de certificado/segurança

Para desenvolvimento/teste local:

1. Abra **certmgr.msc** (Gerenciador de Certificados)
2. Vá em **Certificados - Usuário Atual → Pessoas Confiáveis → Certificados**
3. Se o certificado "VortexExcelAddIn_TemporaryKey" não estiver lá:
   - Clique com o botão direito em "Certificados"
   - Selecione **Todas as Tarefas → Importar**
   - Navegue até `VortexExcelAddIn_TemporaryKey.pfx`
   - Complete a importação

## Desenvolvimento

### Compilando o projeto

```bash
# Via MSBuild
msbuild VortexExcelAddIn\VortexExcelAddIn.csproj /p:Configuration=Release

# Via Visual Studio
# Abra o projeto e pressione Ctrl+Shift+B
```

### Estrutura do Projeto

```
VortexExcelAddIn/
├── Models/              # Modelos de dados
├── Services/            # Serviços (InfluxDB, Excel, Logging, Config)
├── ViewModels/          # ViewModels MVVM
├── Views/               # Interfaces WPF (XAML)
├── Resources/           # Recursos e estilos
├── Properties/          # Configurações do projeto
└── ThisAddIn.cs         # Ponto de entrada do add-in
```

### Tecnologias Utilizadas

- **.NET Framework 4.8**
- **VSTO (Visual Studio Tools for Office)**
- **WPF (Windows Presentation Foundation)**
- **CommunityToolkit.Mvvm** - MVVM toolkit
- **HttpClient** - Cliente HTTP para InfluxDB REST API
- **Newtonsoft.Json** - Serialização JSON
- **NLog** - Sistema de logging
- **CsvHelper** - Exportação CSV

## Suporte

Para problemas, sugestões ou dúvidas:

1. Verifique os logs em `%AppData%/VortexExcelAddIn/logs/`
2. Consulte a seção "Solução de Problemas" acima
3. Abra uma issue no repositório do projeto

## Licença

[Adicione informações de licença aqui]

## Autores

[Adicione informações dos autores aqui]
