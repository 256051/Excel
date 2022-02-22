using System.Collections.Generic;
using System.Data;
using OfficeOpenXml;

namespace YunMa.Excel.Exporter.Base
{
    internal interface IExport<T> where T : class, new()
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        ExcelPackage ExportExcel(DataTable dataItems);

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <returns>文件</returns>
        ExcelPackage ExportExcel(ICollection<T> dataItems);
    }
}