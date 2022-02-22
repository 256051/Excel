using YunMa.Excel.Core.Models;
using YunMa.Excel.Exporter.Base.Filter;

namespace YunMa.Excel.Tests.ExporterTests.Filters
{
    public class TestExporterHeaderFilter1 : IExporterHeaderFilter
    {
        public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
        {
            if (exporterHeaderInfo.DisplayName.Equals("名称"))
            {
                exporterHeaderInfo.DisplayName = "Name";
            }
            return exporterHeaderInfo;
        }
    }
}