namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Configuração de tabela/schema para bancos relacionais.
    /// Permite ao usuário configurar qual tabela e schema usar, além do mapeamento de colunas.
    /// </summary>
    public class TableSchema
    {
        /// <summary>
        /// Nome do schema (ex: "public" para PostgreSQL, "dbo" para SQL Server).
        /// </summary>
        public string SchemaName { get; set; }

        /// <summary>
        /// Nome da tabela.
        /// </summary>
        public string TableName { get; set; }

        /// <summary>
        /// Mapeamento de colunas da tabela para o modelo VortexDataPoint.
        /// </summary>
        public ColumnMapping ColumnMapping { get; set; }

        public TableSchema()
        {
            SchemaName = "public"; // padrão PostgreSQL
            TableName = string.Empty;
            ColumnMapping = new ColumnMapping();
        }
    }

    /// <summary>
    /// Mapeamento de colunas da tabela para propriedades do VortexDataPoint.
    /// Permite flexibilidade para trabalhar com diferentes estruturas de tabelas.
    /// </summary>
    public class ColumnMapping
    {
        /// <summary>
        /// Nome da coluna que contém timestamp/data-hora.
        /// </summary>
        public string TimeColumn { get; set; }

        /// <summary>
        /// Nome da coluna que contém o valor.
        /// </summary>
        public string ValueColumn { get; set; }

        /// <summary>
        /// Nome da coluna que contém o ID do coletor.
        /// </summary>
        public string ColetorIdColumn { get; set; }

        /// <summary>
        /// Nome da coluna que contém o ID do gateway.
        /// </summary>
        public string GatewayIdColumn { get; set; }

        /// <summary>
        /// Nome da coluna que contém o ID do equipamento.
        /// </summary>
        public string EquipmentIdColumn { get; set; }

        /// <summary>
        /// Nome da coluna que contém o ID da tag.
        /// </summary>
        public string TagIdColumn { get; set; }

        public ColumnMapping()
        {
            // Valores padrão baseados na estrutura atual
            TimeColumn = "timestamp";
            ValueColumn = "valor";
            ColetorIdColumn = "coletor_id";
            GatewayIdColumn = "gateway_id";
            EquipmentIdColumn = "equipment_id";
            TagIdColumn = "tag_id";
        }
    }
}
