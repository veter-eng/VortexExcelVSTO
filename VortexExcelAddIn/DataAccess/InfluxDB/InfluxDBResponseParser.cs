using System;
using System.Collections.Generic;
using System.Linq;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.DataAccess.InfluxDB
{
    /// <summary>
    /// Parser de respostas CSV do InfluxDB.
    /// Implementa SRP (Single Responsibility Principle) - responsabilidade única de fazer parsing de CSV.
    /// Extraído do InfluxDBService para separação de responsabilidades.
    /// </summary>
    public class InfluxDBResponseParser
    {
        /// <summary>
        /// Faz parse da resposta CSV do InfluxDB para lista de VortexDataPoint.
        /// Extraído do InfluxDBService linhas 306-496.
        /// </summary>
        public List<VortexDataPoint> Parse(string csvResponse)
        {
            var dataPoints = new List<VortexDataPoint>();

            try
            {
                var lines = csvResponse.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                LoggingService.Info($"Total de linhas na resposta CSV: {lines.Length}");

                if (lines.Length < 2)
                {
                    LoggingService.Warn("Resposta CSV vazia ou sem dados suficientes");
                    return dataPoints;
                }

                // Primeira linha é sempre o header (com ou sem #datatype antes)
                string[] headers = null;
                int dataStartIndex = 1;

                // Verificar se tem #datatype na primeira linha
                if (lines[0].StartsWith("#datatype"))
                {
                    // Formato com #datatype: próxima linha é o header
                    if (lines.Length > 2)
                    {
                        headers = lines[1].Split(',');
                        dataStartIndex = 2;
                    }
                }
                else
                {
                    // Formato direto: primeira linha já é o header
                    headers = lines[0].Split(',');
                    dataStartIndex = 1;
                }

                if (headers == null || headers.Length == 0)
                {
                    LoggingService.Warn("Nenhum header encontrado na resposta CSV");
                    return dataPoints;
                }

                LoggingService.Info($"Total de headers: {headers.Length}");
                LoggingService.Info($"Headers encontrados: [{string.Join("] [", headers)}]");
                LoggingService.Info($"Dados começam na linha: {dataStartIndex}");

                // Indices das colunas
                var timeIndex = Array.IndexOf(headers, "_time");
                var valueIndex = Array.IndexOf(headers, "_value");
                var coletorIndex = Array.IndexOf(headers, "coletor_id");
                var gatewayIndex = Array.IndexOf(headers, "gateway_id");
                var equipmentIndex = Array.IndexOf(headers, "equipment_id");
                var tagIndex = Array.IndexOf(headers, "tag_id");

                LoggingService.Info($"Índices das colunas - Time:{timeIndex}, Value:{valueIndex}, Coletor:{coletorIndex}, Gateway:{gatewayIndex}, Equipment:{equipmentIndex}, Tag:{tagIndex}");

                // Log das primeiras linhas de dados para debug
                if (lines.Length > dataStartIndex)
                {
                    for (int debugIdx = dataStartIndex; debugIdx < Math.Min(dataStartIndex + 3, lines.Length); debugIdx++)
                    {
                        var debugValues = lines[debugIdx].Split(',');
                        LoggingService.Info($"Linha {debugIdx} tem {debugValues.Length} valores: [{string.Join("] [", debugValues)}]");
                    }
                }

                // Parse dos dados
                int dataLineCount = 0;

                for (int i = dataStartIndex; i < lines.Length; i++)
                {
                    var line = lines[i];

                    // Pular apenas linhas completamente vazias
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    var values = line.Split(',');

                    // Detectar linhas de header repetidas (quando há múltiplas tables)
                    // Headers sempre começam com vírgula vazia e contêm nomes de colunas do InfluxDB
                    bool isHeaderLine = false;

                    // Verificar se parece ser um header:
                    // 1. Começa com vírgula (primeira coluna vazia)
                    // 2. Contém palavras-chave de colunas do InfluxDB
                    if (line.StartsWith(",") && values.Length > 5)
                    {
                        // Verificar se contém colunas típicas do InfluxDB
                        var lineUpper = line.ToLower();
                        if (lineUpper.Contains("_time") && lineUpper.Contains("_value") &&
                            (lineUpper.Contains("_field") || lineUpper.Contains("_measurement")))
                        {
                            isHeaderLine = true;
                        }
                    }

                    if (isHeaderLine)
                    {
                        // Nova table detectada - atualizar headers
                        headers = values;

                        // Recalcular índices das colunas
                        timeIndex = Array.IndexOf(headers, "_time");
                        valueIndex = Array.IndexOf(headers, "_value");
                        coletorIndex = Array.IndexOf(headers, "coletor_id");
                        gatewayIndex = Array.IndexOf(headers, "gateway_id");
                        equipmentIndex = Array.IndexOf(headers, "equipment_id");
                        tagIndex = Array.IndexOf(headers, "tag_id");

                        LoggingService.Info($"***** NOVA TABLE DETECTADA na linha {i} ({dataLineCount} registros até agora) *****");
                        LoggingService.Info($"Headers: [{string.Join("] [", headers)}]");
                        LoggingService.Info($"Novos índices - Time:{timeIndex}, Value:{valueIndex}, Coletor:{coletorIndex}, Gateway:{gatewayIndex}, Equipment:{equipmentIndex}, Tag:{tagIndex}");
                        continue;
                    }

                    // Processar TODAS as linhas de dados, sem filtrar por número de colunas
                    // Se a linha tem pelo menos 1 valor, processar
                    if (values.Length == 0)
                    {
                        continue;
                    }

                    dataLineCount++;

                    // Log detalhado ao redor da área problemática (linha ~4000 de dados)
                    bool isNearProblemArea = (dataLineCount >= 3995 && dataLineCount <= 4005);

                    if (isNearProblemArea)
                    {
                        LoggingService.Info($">>> Linha CSV {i}, DataPoint #{dataLineCount}: {values.Length} valores");
                        LoggingService.Info($">>> Linha bruta: {line.Substring(0, Math.Min(200, line.Length))}");
                        LoggingService.Info($">>> Índices atuais - Time:{timeIndex}, Value:{valueIndex}, Coletor:{coletorIndex}, Gateway:{gatewayIndex}, Equipment:{equipmentIndex}, Tag:{tagIndex}");
                    }

                    var dataPoint = new VortexDataPoint
                    {
                        Time = timeIndex >= 0 && timeIndex < values.Length ? ParseTimestamp(CleanCsvValue(values[timeIndex])) : DateTime.UtcNow,
                        Valor = valueIndex >= 0 && valueIndex < values.Length ? CleanCsvValue(values[valueIndex]) : string.Empty,
                        ColetorId = coletorIndex >= 0 && coletorIndex < values.Length ? CleanCsvValue(values[coletorIndex]) : string.Empty,
                        GatewayId = gatewayIndex >= 0 && gatewayIndex < values.Length ? CleanCsvValue(values[gatewayIndex]) : string.Empty,
                        EquipmentId = equipmentIndex >= 0 && equipmentIndex < values.Length ? CleanCsvValue(values[equipmentIndex]) : string.Empty,
                        TagId = tagIndex >= 0 && tagIndex < values.Length ? CleanCsvValue(values[tagIndex]) : string.Empty
                    };

                    // REMOVIDO: Validação de IDs numéricos estava filtrando registros válidos
                    // Isso causava retorno de menos de 1000 registros mesmo com limit=1000
                    // Os IDs são strings no modelo e devem aceitar qualquer valor
                    //
                    // Código anterior (comentado para referência):
                    // bool isValidColetorId = string.IsNullOrWhiteSpace(dataPoint.ColetorId) ||
                    //                        int.TryParse(dataPoint.ColetorId, out _);
                    // bool isValidGatewayId = string.IsNullOrWhiteSpace(dataPoint.GatewayId) ||
                    //                        int.TryParse(dataPoint.GatewayId, out _);
                    // bool isValidEquipmentId = string.IsNullOrWhiteSpace(dataPoint.EquipmentId) ||
                    //                          int.TryParse(dataPoint.EquipmentId, out _);
                    // bool isValidTagId = string.IsNullOrWhiteSpace(dataPoint.TagId) ||
                    //                    int.TryParse(dataPoint.TagId, out _);
                    //
                    // if (!isValidColetorId || !isValidGatewayId || !isValidEquipmentId || !isValidTagId)
                    // {
                    //     invalidIdCount++;
                    //     if (invalidIdCount <= 10)
                    //     {
                    //         LoggingService.Debug($"Linha {i} ignorada - IDs não numéricos: Coletor=[{dataPoint.ColetorId}], Gateway=[{dataPoint.GatewayId}], Equipment=[{dataPoint.EquipmentId}], Tag=[{dataPoint.TagId}]");
                    //     }
                    //     continue;
                    // }

                    dataPoints.Add(dataPoint);

                    // Log detalhado dos primeiros registros para debug
                    if (dataLineCount <= 5 || isNearProblemArea)
                    {
                        LoggingService.Info($"DataPoint {dataLineCount} (linha CSV {i}): Time={dataPoint.Time:HH:mm:ss}, Valor=[{dataPoint.Valor}], Coletor=[{dataPoint.ColetorId}], Gateway=[{dataPoint.GatewayId}], Equipment=[{dataPoint.EquipmentId}], Tag=[{dataPoint.TagId}]");
                    }
                }

                LoggingService.Info($"Total de registros parseados: {dataLineCount}");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao fazer parse da resposta Flux", ex);
            }

            return dataPoints;
        }

        /// <summary>
        /// Parse de valores distintos da resposta CSV.
        /// Extraído do InfluxDBService linhas 501-547.
        /// </summary>
        public List<string> ParseDistinctValues(string csvResponse, string columnName)
        {
            var values = new List<string>();

            try
            {
                var lines = csvResponse.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                var dataStartIndex = -1;
                string[] headers = null;

                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith("#datatype"))
                    {
                        if (i + 1 < lines.Length)
                        {
                            headers = lines[i + 1].Split(',');
                            dataStartIndex = i + 2;
                            break;
                        }
                    }
                }

                if (dataStartIndex == -1 || headers == null)
                    return values;

                var columnIndex = Array.IndexOf(headers, columnName);
                if (columnIndex == -1)
                    return values;

                for (int i = dataStartIndex; i < lines.Length; i++)
                {
                    var cols = lines[i].Split(',');
                    if (columnIndex < cols.Length && !string.IsNullOrEmpty(cols[columnIndex]))
                    {
                        values.Add(cols[columnIndex]);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao fazer parse de valores distintos para {columnName}", ex);
            }

            return values.Distinct().ToList();
        }

        /// <summary>
        /// Parse de timestamp RFC3339.
        /// Extraído do InfluxDBService linhas 552-558.
        /// </summary>
        private DateTime ParseTimestamp(string timestamp)
        {
            if (DateTime.TryParse(timestamp, out var result))
                return result.ToUniversalTime();

            return DateTime.UtcNow;
        }

        /// <summary>
        /// Remove aspas duplas do início e fim de um valor CSV.
        /// Extraído do InfluxDBService linhas 563-580.
        /// </summary>
        private string CleanCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            // Remove aspas duplas do início e fim
            value = value.Trim();
            if (value.StartsWith("\"") && value.EndsWith("\"") && value.Length > 1)
            {
                value = value.Substring(1, value.Length - 2);
            }
            else if (value.StartsWith("\""))
            {
                value = value.Substring(1);
            }

            return value;
        }
    }
}
