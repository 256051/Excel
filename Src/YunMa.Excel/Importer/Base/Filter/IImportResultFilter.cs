using YunMa.Excel.Core.Models;
using YunMa.Excel.Filters;

namespace YunMa.Excel.Importer.Base.Filter
{
    public interface IImportResultFilter : IFilter
    {
        /// <summary>
        /// 处理导入结果
        /// 比如对错误信息进行多语言转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        ImportResult<T> Filter<T>(ImportResult<T> importResult) where T : class, new();
    }
}