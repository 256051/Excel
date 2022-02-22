using System;
using System.Threading.Tasks;
using YunMa.Excel.Core.Models;

namespace YunMa.Excel.Importer.Base
{
    internal interface IImportMultipleSheet : IDisposable
    {
        Task<ImportResult<object>> Import(string sheetName, int sheetIndex, Type importDataType,
            bool isSaveLabelingError = true);

        Task<byte[]> GenerateTemplateByte();
    }
}