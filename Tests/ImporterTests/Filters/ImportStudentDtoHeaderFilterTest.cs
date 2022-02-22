using System.Collections.Generic;
using YunMa.Excel.Core.Models;
using YunMa.Excel.Importer.Base.Filter;

namespace YunMa.Excel.Tests.ImporterTests.Filters
{
    /// <summary>
    /// 导入列头筛选器测试
    /// 1）测试修改列头
    /// 2）测试修改值映射
    /// </summary>
    public class ImportStudentDtoHeaderFilterTest : IImportHeaderFilter
    {
        public List<ImporterHeaderInfo> Filter(List<ImporterHeaderInfo> importerHeaderInfos)
        {
            foreach (var item in importerHeaderInfos)
            {
                if (item.PropertyName == "Gender")
                {
                    item.MappingValues = new Dictionary<string, dynamic>()
                    {
                        {"男0",0 },
                        {"女1",1 }
                    };
                }
            }
            return importerHeaderInfos;
        }
    }

}