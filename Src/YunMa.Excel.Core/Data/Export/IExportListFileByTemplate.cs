using System.Collections.Generic;
using System.Threading.Tasks;
using YunMa.Excel.Core.Models;

namespace YunMa.Excel.Core.Data.Export
{
    
    /// <summary>
    /// 根据模板导出列表文件
    /// </summary>
    public interface IExportListFileByTemplate
    {
        /// <summary>
        ///     根据模板导出列表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportListByTemplate<T>(string fileName, ICollection<T> dataItems,
            string htmlTemplate = null) where T : class;
    }
}