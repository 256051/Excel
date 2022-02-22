using System;
using System.Threading.Tasks;
using YunMa.Excel.Core.Models;

namespace YunMa.Excel.Importer.Base
{
    internal interface IImport<T> : IDisposable where T : class, new()
    {
        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        Task<ImportResult<T>> ImportExcel(string filePath = null);
    }
}