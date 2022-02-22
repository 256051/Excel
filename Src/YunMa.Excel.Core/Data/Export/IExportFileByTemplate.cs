using System;
using System.Threading.Tasks;
using YunMa.Excel.Core.Models;

namespace YunMa.Excel.Core.Data.Export
{
    public interface IExportFileByTemplate
    {
        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        /// <param name="template">HTML模板或模板路径</param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportByTemplate<T>(string fileName, T data,
            string template) where T : class;

        /// <summary>
        ///     根据模板导出到载荷
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template">HTML模板或模板路径</param>
        /// <returns></returns>
        Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class;
        // /// <summary>
        // ///		根据模板导出
        // /// </summary>
        // /// <param name="data"></param>
        // /// <param name="template"></param>
        // /// <param name="type"></param>
        // /// <returns></returns>
        // Task<byte[]> ExportBytesByTemplate(object data, string template, Type type);
    }
}