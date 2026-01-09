using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Office.Interop.Excel;
using VortexExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço para manipulação de dados no Excel usando Interop
    /// Port do ExcelService.ts
    /// </summary>
    public static class ExcelService
    {
        // Cores do tema Vortex (RGB)
        private static readonly Color PrimaryColor = Color.FromArgb(68, 114, 196);  // #4472C4
        private static readonly Color WhiteColor = Color.White;

        /// <summary>
        /// Exporta dados para uma planilha (ativa ou nova)
        /// </summary>
        /// <param name="data">Lista de dados para exportar</param>
        /// <param name="sheetName">Nome da planilha (opcional)</param>
        /// <param name="databaseType">Tipo de banco de dados para ajustar headers</param>
        public static void ExportToSheet(List<VortexDataPoint> data, string sheetName = null, Domain.Models.DatabaseType? databaseType = null)
        {
            if (data == null || data.Count == 0)
            {
                throw new ArgumentException("Dados vazios para exportar", nameof(data));
            }

            Worksheet sheet = null;
            Range range = null;
            Range headerRange = null;
            ListObject table = null;

            try
            {
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;

                if (workbook == null)
                {
                    throw new InvalidOperationException("Nenhum workbook ativo");
                }

                // Criar ou obter planilha
                if (!string.IsNullOrEmpty(sheetName))
                {
                    // Tentar encontrar planilha existente
                    sheet = FindWorksheet(workbook, sheetName);

                    if (sheet == null)
                    {
                        // Criar nova planilha
                        sheet = (Worksheet)workbook.Worksheets.Add();
                        sheet.Name = sheetName;
                        LoggingService.Info($"Nova planilha criada: {sheetName}");
                    }
                    else
                    {
                        LoggingService.Info($"Usando planilha existente: {sheetName}");
                    }
                }
                else
                {
                    sheet = (Worksheet)app.ActiveSheet;
                }

                // Preparar array 2D (mais eficiente que célula por célula)
                // VortexIO tem 5 colunas (sem Coletor ID), Vortex Historian e Vortex Historian API têm 6 colunas
                bool isVortexIO = databaseType == Domain.Models.DatabaseType.VortexAPI;
                int columnCount = isVortexIO ? 5 : 6;
                object[,] values = new object[data.Count + 1, columnCount];

                // Cabeçalhos - Ajustados para VortexIO (dados_airflow) vs Vortex Historian/Historian API (dados_rabbitmq)
                if (isVortexIO)
                {
                    // VortexIO: Timestamp, Campo, Tipo de Agregação, Tag ID, Valor
                    values[0, 0] = "Timestamp";
                    values[0, 1] = "Campo";  // _field (avg_valor, sum_valor, etc.)
                    values[0, 2] = "Tipo de Agregação";  // aggregation_type (average_60m, total_60m, etc.)
                    values[0, 3] = "Tag ID";
                    values[0, 4] = "Valor";
                }
                else
                {
                    // Vortex Historian: Timestamp, Coletor ID, Gateway ID, Equipment ID, Tag ID, Valor
                    values[0, 0] = "Timestamp";
                    values[0, 1] = "Coletor ID";
                    values[0, 2] = "Gateway ID";
                    values[0, 3] = "Equipment ID";
                    values[0, 4] = "Tag ID";
                    values[0, 5] = "Valor";
                }

                // Dados
                for (int i = 0; i < data.Count; i++)
                {
                    if (isVortexIO)
                    {
                        // VortexIO: Skip ColetorId (not used)
                        values[i + 1, 0] = data[i].Time;  // DateTime object (not string)
                        values[i + 1, 1] = data[i].GatewayId;  // Campo (_field)
                        values[i + 1, 2] = data[i].EquipmentId;  // Tipo de Agregação
                        values[i + 1, 3] = data[i].TagId;
                        values[i + 1, 4] = data[i].Valor;
                    }
                    else
                    {
                        // Vortex Historian: All fields
                        values[i + 1, 0] = data[i].Time;  // DateTime object (not string)
                        values[i + 1, 1] = data[i].ColetorId;
                        values[i + 1, 2] = data[i].GatewayId;
                        values[i + 1, 3] = data[i].EquipmentId;
                        values[i + 1, 4] = data[i].TagId;
                        values[i + 1, 5] = data[i].Valor;
                    }
                }

                // Escrever dados em batch (muito mais rápido)
                range = (Excel.Range)sheet.Cells[1, 1];
                range = range.Resize[data.Count + 1, columnCount];
                range.Value2 = values;

                // Formatar coluna de timestamp (primeira coluna) com formato brasileiro
                var timestampColumn = (Excel.Range)sheet.Cells[2, 1];  // Começar na linha 2 (pular cabeçalho)
                timestampColumn = timestampColumn.Resize[data.Count, 1];
                timestampColumn.NumberFormat = "dd/mm/yyyy hh:mm:ss";

                // Formatar cabeçalhos
                headerRange = (Range)sheet.Cells[1, 1];
                headerRange = headerRange.Resize[1, columnCount];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = ColorTranslator.ToOle(PrimaryColor);
                headerRange.Font.Color = ColorTranslator.ToOle(WhiteColor);

                // Auto-fit colunas
                range.Columns.AutoFit();

                // Criar tabela Excel
                table = sheet.ListObjects.Add(
                    XlListObjectSourceType.xlSrcRange,
                    range,
                    Type.Missing,
                    XlYesNoGuess.xlYes,
                    Type.Missing
                );

                table.Name = $"VortexData_{DateTime.Now.Ticks}";
                table.TableStyle = "TableStyleMedium2";

                LoggingService.Info($"Dados exportados com sucesso: {data.Count} registros");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao exportar para planilha", ex);
                throw new Exception("Não foi possível exportar dados para a planilha", ex);
            }
            finally
            {
                // Liberar objetos COM para evitar memory leaks
                if (table != null) Marshal.ReleaseComObject(table);
                if (headerRange != null) Marshal.ReleaseComObject(headerRange);
                if (range != null) Marshal.ReleaseComObject(range);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        /// <summary>
        /// Atualiza dados em uma tabela existente
        /// </summary>
        public static void UpdateSheet(List<VortexDataPoint> data, string tableName)
        {
            if (data == null || data.Count == 0)
            {
                throw new ArgumentException("Dados vazios para atualizar", nameof(data));
            }

            if (string.IsNullOrEmpty(tableName))
            {
                throw new ArgumentException("Nome da tabela não pode ser vazio", nameof(tableName));
            }

            ListObject table = null;
            Range dataRange = null;

            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    throw new InvalidOperationException("Nenhum workbook ativo");
                }

                // Encontrar tabela
                table = FindTable(workbook, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Tabela '{tableName}' não encontrada");
                }

                // Limpar dados existentes (manter cabeçalhos)
                if (table.DataBodyRange != null)
                {
                    table.DataBodyRange.Delete(XlDeleteShiftDirection.xlShiftUp);
                }

                // Preparar novos dados
                object[,] values = new object[data.Count, 6];
                for (int i = 0; i < data.Count; i++)
                {
                    values[i, 0] = data[i].Time;  // DateTime object (not string)
                    values[i, 1] = data[i].ColetorId;
                    values[i, 2] = data[i].GatewayId;
                    values[i, 3] = data[i].EquipmentId;
                    values[i, 4] = data[i].TagId;
                    values[i, 5] = data[i].Valor;
                }

                // Adicionar novas linhas
                dataRange = table.ListRows.Add().Range;
                dataRange = dataRange.Resize[data.Count, 6];
                dataRange.Value2 = values;

                // Formatar coluna de timestamp com formato brasileiro
                var timestampColumn = (Excel.Range)dataRange.Cells[1, 1];
                timestampColumn = timestampColumn.Resize[data.Count, 1];
                timestampColumn.NumberFormat = "dd/mm/yyyy hh:mm:ss";

                LoggingService.Info($"Tabela '{tableName}' atualizada: {data.Count} registros");
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao atualizar tabela '{tableName}'", ex);
                throw new Exception("Não foi possível atualizar a planilha", ex);
            }
            finally
            {
                if (dataRange != null) Marshal.ReleaseComObject(dataRange);
                if (table != null) Marshal.ReleaseComObject(table);
            }
        }

        /// <summary>
        /// Cria um gráfico a partir dos dados
        /// </summary>
        public static void CreateChart(List<VortexDataPoint> data, XlChartType chartType = XlChartType.xlLine)
        {
            if (data == null || data.Count == 0)
            {
                throw new ArgumentException("Dados vazios para criar gráfico", nameof(data));
            }

            Worksheet sheet = null;
            Range dataRange = null;
            Chart chart = null;

            try
            {
                var app = Globals.ThisAddIn.Application;
                sheet = (Worksheet)app.ActiveSheet;

                // Preparar dados para o gráfico (somente Timestamp e Valor)
                object[,] values = new object[data.Count + 1, 2];
                values[0, 0] = "Timestamp";
                values[0, 1] = "Valor";

                for (int i = 0; i < data.Count; i++)
                {
                    values[i + 1, 0] = data[i].Time;  // DateTime object (not string)

                    // Tentar converter valor para número
                    if (double.TryParse(data[i].Valor, out double valor))
                    {
                        values[i + 1, 1] = valor;
                    }
                    else
                    {
                        values[i + 1, 1] = 0;
                    }
                }

                // Criar range temporário para dados do gráfico (coluna K e L)
                dataRange = (Range)sheet.Cells[1, 11]; // Coluna K
                dataRange = dataRange.Resize[data.Count + 1, 2];
                dataRange.Value2 = values;

                // Criar gráfico
                var charts = sheet.ChartObjects() as ChartObjects;
                var chartObject = charts.Add(50, 300, 500, 300);
                chart = chartObject.Chart;

                chart.SetSourceData(dataRange);
                chart.ChartType = chartType;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Dados Vortex";
                chart.HasLegend = true;
                chart.Legend.Position = XlLegendPosition.xlLegendPositionRight;

                LoggingService.Info($"Gráfico criado com sucesso: {data.Count} pontos");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao criar gráfico", ex);
                throw new Exception("Não foi possível criar o gráfico", ex);
            }
            finally
            {
                if (chart != null) Marshal.ReleaseComObject(chart);
                if (dataRange != null) Marshal.ReleaseComObject(dataRange);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        /// <summary>
        /// Exporta dados para string CSV
        /// </summary>
        public static string ExportToCSV(List<VortexDataPoint> data)
        {
            if (data == null || data.Count == 0)
            {
                return string.Empty;
            }

            try
            {
                var sb = new StringBuilder();
                using (var writer = new StringWriter(sb))
                using (var csv = new CsvWriter(writer, new CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture)))
                {
                    // Escrever headers
                    csv.WriteField("Timestamp");
                    csv.WriteField("Coletor ID");
                    csv.WriteField("Gateway ID");
                    csv.WriteField("Equipment ID");
                    csv.WriteField("Tag ID");
                    csv.WriteField("Valor");
                    csv.NextRecord();

                    // Escrever dados
                    foreach (var point in data)
                    {
                        csv.WriteField(point.Time.ToString("yyyy-MM-dd HH:mm:ss"));
                        csv.WriteField(point.ColetorId);
                        csv.WriteField(point.GatewayId);
                        csv.WriteField(point.EquipmentId);
                        csv.WriteField(point.TagId);
                        csv.WriteField(point.Valor);
                        csv.NextRecord();
                    }
                }

                LoggingService.Info($"CSV gerado: {data.Count} registros");
                return sb.ToString();
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao gerar CSV", ex);
                throw new Exception("Não foi possível exportar para CSV", ex);
            }
        }

        /// <summary>
        /// Salva dados em arquivo CSV usando SaveFileDialog
        /// </summary>
        public static void DownloadCsv(List<VortexDataPoint> data, string defaultFileName = "vortex_data.csv")
        {
            if (data == null || data.Count == 0)
            {
                MessageBox.Show("Não há dados para exportar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = defaultFileName;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        var csvContent = ExportToCSV(data);
                        File.WriteAllText(saveFileDialog.FileName, csvContent, Encoding.UTF8);

                        LoggingService.Info($"CSV salvo em: {saveFileDialog.FileName}");
                        MessageBox.Show($"Arquivo salvo com sucesso em:\n{saveFileDialog.FileName}",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao salvar CSV", ex);
                MessageBox.Show($"Erro ao salvar arquivo:\n{ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Limpa a planilha ativa
        /// </summary>
        public static void ClearActiveSheet()
        {
            Worksheet sheet = null;
            Range usedRange = null;

            try
            {
                var app = Globals.ThisAddIn.Application;
                sheet = (Worksheet)app.ActiveSheet;
                usedRange = sheet.UsedRange;

                if (usedRange != null)
                {
                    usedRange.Clear();
                    LoggingService.Info("Planilha ativa limpa");
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao limpar planilha", ex);
                throw new Exception("Não foi possível limpar a planilha", ex);
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        #region Helper Methods

        /// <summary>
        /// Encontra uma planilha pelo nome
        /// </summary>
        private static Worksheet FindWorksheet(Workbook workbook, string name)
        {
            try
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        return sheet;
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Encontra uma tabela pelo nome em todas as planilhas
        /// </summary>
        private static ListObject FindTable(Workbook workbook, string tableName)
        {
            try
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        if (table.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            return table;
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}
