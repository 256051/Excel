using System.Linq;
using YunMa.Excel.Exporter.Base.Writerable;

namespace YunMa.Excel.Exporter.Base.Utility
{
    internal class TemplateTableInfo
    {
        /// <summary>
        /// 表格数据对象Key
        /// </summary>
        public string TableKey { get; set; }

        /// <summary>
        /// 原始开始行
        /// </summary>
        public int RawRowStart { get; set; }

        /// <summary>
        /// 新开始行
        /// </summary>
        public int NewRowStart { get; set; }

        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; set; }

        /// <summary>
        /// 写入器
        /// </summary>
        public IGrouping<string, IWriter> Writers { get; set; }
    }
}