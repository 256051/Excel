using OfficeOpenXml.Table;
using YunMa.Excel.Core.Attributes.Export;
using YunMa.Excel.Core.Enums;

namespace YunMa.Excel.Exporter.Base.Attributes
{
    /// <summary>
    ///     Excel导出特性
    /// </summary>
    public class ExcelExporterAttribute : ExporterAttribute
    {
        /// <summary>
        ///  输出类型
        /// </summary>
        public ExcelOutputTypes ExcelOutputType { get; set; } = ExcelOutputTypes.DataTable;

        /// <summary>
        ///     自动居中(设置后为全局居中显示)
        /// </summary>
        public bool AutoCenter { get; set; }

        /// <summary>
        ///     表头位置
        /// </summary>
        public int HeaderRowIndex { get; set; } = 1;


        /// <summary>
        ///     表格样式风格
        /// </summary>
        public TableStyles TableStyle { get; set; } = TableStyles.None;
    }
}