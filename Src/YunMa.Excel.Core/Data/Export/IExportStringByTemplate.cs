using System.Threading.Tasks;

namespace YunMa.Excel.Core.Data.Export
{
    public interface IExportStringByTemplate
    {
        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        Task<string> ExportByTemplate<T>(T data, string htmlTemplate = null) where T : class;
    }
}