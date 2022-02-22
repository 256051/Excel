using System;
using YunMa.Excel.Core.Attributes.Import;

namespace YunMa.Excel.Importer.Base.Attributes
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ExcelImporterAttribute : ImporterAttribute
    {
        /// <summary>
        ///     指定Sheet名称(获取指定Sheet名称)
        ///     为空则自动获取第一个
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        ///     指定Sheet下标（获取指定Sheet下标）
        /// </summary>
        /// <remarks>
        ///     在.NET Core+包括.NET5框架中下标从0开始，否则从 1 
        /// </remarks>
        public int SheetIndex { get; set; } = 0; 



        /// <summary>
        ///     截止读取的列数（从1开始，如果已设置，则将支持空行以及特殊列）
        /// </summary>
        public int? EndColumnCount { get; set; }

        /// <summary>
        ///     是否标注错误（默认为true）
        /// </summary>
        public bool IsLabelingError { get; set; } = true;

        /// <summary>
        /// Sheet顶部导入描述
        /// </summary>
        public string ImportDescription { get; set; }

        /// <summary>
        /// Sheet顶部导入描述高度(换行可能无法自动设定高度,默认为Excel的默认行高)
        /// </summary>
        public double DescriptionHeight { get; set; } = 13.5;

        /// <summary>
        ///     是否仅导出错误数据
        /// </summary>
        public bool IsOnlyErrorRows { get; set; }

        /// <summary>
        /// 结束符
        /// </summary>
        public string ClosingNote { get; set; }
    }
}
