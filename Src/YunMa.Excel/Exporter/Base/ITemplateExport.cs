using System;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace YunMa.Excel.Exporter.Base
{
    /// <summary>
    /// 模版导出
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface ITemplateExport<in T> : IDisposable where T : class
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="templateFilePath"></param>
        /// <param name="data"></param>
        /// <param name="callback"></param>
        /// <returns></returns>
        Task Export(string templateFilePath, T data, Action<ExcelPackage> callback = null);
    }
}