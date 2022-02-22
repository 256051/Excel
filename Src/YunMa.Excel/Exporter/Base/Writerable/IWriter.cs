namespace YunMa.Excel.Exporter.Base.Writerable
{
    /// <summary>
    /// 写入器
    /// </summary>
    public interface IWriter
    {
        /// <summary>
        /// 原始模板地址
        /// </summary>
        string TplAddress { get; set; }

        /// <summary>
        /// 单元格原始字符串
        /// </summary>
        string CellString { get; set; }

        /// <summary>
        /// 行号
        /// </summary>
        int RowIndex { get; set; }

        /// <summary>
        /// 列号
        /// </summary>
        int ColIndex { get; set; }

        /// <summary>
        /// 写入器类型
        /// </summary>
        WriterTypes WriterType { get; set; }

        /// <summary>
        /// 表格数据对象Key
        /// </summary>
        string TableKey { get; set; }
    }
}