using System.Collections.Generic;
using YunMa.Excel.Core.Models;
using YunMa.Excel.Filters;

namespace YunMa.Excel.Importer.Base.Filter
{
    public interface IImportHeaderFilter : IFilter
    {
        /// <summary>
        /// 处理列头
        /// </summary>
        /// <param name="importerHeaderInfos"></param>
        /// <returns></returns>
        List<ImporterHeaderInfo> Filter(List<ImporterHeaderInfo> importerHeaderInfos);
    }
}