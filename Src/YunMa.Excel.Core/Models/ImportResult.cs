using System;
using System.Collections.Generic;
using System.Linq;

namespace YunMa.Excel.Core.Models
{
    /// <summary>
    ///     导入结果
    /// </summary>
    public sealed class ImportResult<T> where T : class
    {
        /// <summary>
        /// </summary>
        public ImportResult()
        {
            RowErrors = new List<DataRowErrorInfo>();
        }

        /// <summary>
        ///     导入数据
        /// </summary>
        public ICollection<T> Data { get; set; }

        /// <summary>
        ///     验证错误
        /// </summary>
        public IList<DataRowErrorInfo> RowErrors { get; set; }

        /// <summary>
        ///     模板错误
        /// </summary>
        public IList<TemplateErrorInfo> TemplateErrors { get; set; }

        /// <summary>
        ///     导入异常信息
        /// </summary>
        public Exception Exception { get; set; }

        /// <summary>
        ///     是否存在导入错误
        /// </summary>
        public bool HasError => Exception != null ||
                                (TemplateErrors?.Count(p => p.ErrorLevel == ErrorLevels.Error) ?? 0) > 0 ||
                                (RowErrors?.Count ?? 0) > 0;

        /// <summary>
        ///     Imported header list information
        ///     导入的表头列表信息
        ///     https://github.com/dotnetcore/Magicodes.IE/issues/76
        /// </summary>
        public IList<ImporterHeaderInfo> ImporterHeaderInfos { get; set; }
    }
}