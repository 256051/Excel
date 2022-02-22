using OfficeOpenXml.Table;
using YunMa.Excel.Core.Attributes.Export;
using YunMa.Excel.Exporter.Base.Attributes;
using YunMa.Excel.Tests.ExporterTests.Filters;

namespace YunMa.Excel.Tests.ExporterTests.Models
{

    /// <summary>
    /// 导出注入
    /// </summary>
    [ExcelExporter(Name = "测试", TableStyle = TableStyles.Light10, ExporterHeaderFilter = typeof(TestExporterHeaderFilter1))]
    public class ExporterHeaderFilterTestData1
    {
        [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
        public string Text { get; set; }

        [ExporterHeader(DisplayName = "普通文本")] public string Text2 { get; set; }

        [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
        public string Text3 { get; set; }

        [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
        public decimal Number { get; set; }

        [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
        public string Name { get; set; }
    }
}