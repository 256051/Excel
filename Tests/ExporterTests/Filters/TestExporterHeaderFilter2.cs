using YunMa.Excel.Core.Models;
using YunMa.Excel.Exporter.Base.Filter;

namespace YunMa.Excel.Tests.ExporterTests.Filters
{
    public class TestExporterHeaderFilter2 : IExporterHeaderFilter
    {
        /// <summary>
        /// 表头筛选器（修改忽略列）
        /// </summary>
        /// <param name="exporterHeaderInfo"></param>
        /// <returns></returns>
        public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
        {
            if (exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore)
            {
                exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore = false;
            }
            return exporterHeaderInfo;
        }
    }
}