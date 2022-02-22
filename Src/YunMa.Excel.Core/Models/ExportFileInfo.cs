namespace YunMa.Excel.Core.Models
{
    /// <summary>
    /// 导出文件信息
    /// </summary>
    public class ExportFileInfo
    {
        /// <summary>
        ///
        /// </summary>
        public ExportFileInfo()
        {
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileType"></param>
        public ExportFileInfo(string fileName, string fileType)
        {
            FileName = fileName;
            FileType = fileType;
        }

        /// <summary>
        /// 文件名（路径）
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件Mine类型
        /// </summary>
        public string FileType { get; set; }
    }
}