namespace YunMa.Excel.Exporter.Base.Writerable
{
    /// <summary>
    /// 写入器
    /// </summary>
    public class Writer : IWriter
    {
        /// <summary>
        /// 地址
        /// </summary>
        public string TplAddress { get; set; }

        /// <summary>
        /// 单元格原始字符串
        /// </summary>
        public string CellString { get; set; }

        /// <summary>
        /// 写入的字符串
        /// </summary>
        public string WriteString { get; set; }

        /// <summary>
        /// 写入器类型
        /// </summary>
        public WriterTypes WriterType { get; set; }

        /// <summary>
        /// 表格数据对象Key
        /// </summary>
        public string TableKey { get; set; }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 列号
        /// </summary>
        public int ColIndex { get; set; }
    }
}