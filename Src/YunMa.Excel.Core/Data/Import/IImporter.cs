using System.IO;
using System.Threading.Tasks;
using YunMa.Excel.Core.Models;

namespace YunMa.Excel.Core.Data.Import
{
    public interface IImporter
    {
        /// <summary> 
        ///     生成导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<ExportFileInfo> GenerateTemplate<T>(string fileName) where T : class, new();

        /// <summary>
        ///     生成导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>二进制字节</returns>
        Task<byte[]> GenerateTemplateBytes<T>() where T : class, new();

        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="labelingFilePath">标注文件路径</param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(string filePath, string labelingFilePath = null) where T : class, new();

        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(Stream stream) where T : class, new();



        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="labelStream"></param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(Stream stream, Stream labelStream) where T : class, new();


        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="labelStream"></param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(string filePath, Stream labelStream) where T : class, new();
    }
}