using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using YunMa.Excel.Core;
using YunMa.Excel.Core.Extensions;
using YunMa.Excel.Core.Models;
using YunMa.Excel.Importer.Base.Attributes;
using YunMa.Excel.Importer.Base.Filter;
using YunMa.Excel.Importer.Base.Impl;

namespace YunMa.Excel.Importer
{
    public class ExcelImporter : IExcelImporter
    {
        private readonly IEnumerable<IImportHeaderFilter> _importHeaderFilters;
        private readonly IEnumerable<IImportResultFilter> _importResultFilters;


        public ExcelImporter(IEnumerable<IImportHeaderFilter> importHeaderFilters, IEnumerable<IImportResultFilter> importResultFilters)
        {
            _importHeaderFilters = importHeaderFilters;
            _importResultFilters = importResultFilters;
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">文件名必须填写! - fileName</exception>
        public Task<ExportFileInfo> GenerateTemplate<T>(string fileName) where T : class, new()
        {
            fileName.CheckExcelFileName();
            var isMultipleSheetType = false;
            var tableType = typeof(T);
            List<PropertyInfo> sheetPropertyList = new List<PropertyInfo>();
            var sheetProperties = tableType.GetProperties();

            for (var i = 0; i < sheetProperties.Length; i++)
            {
                var sheetProperty = sheetProperties[i];
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(importerAttribute.SheetName))
                {
                    isMultipleSheetType = true;
                    sheetPropertyList.Add(sheetProperty);
                }
            }

            if (isMultipleSheetType)
            {
                using var importer = new ImportMultipleSheet(sheetPropertyList);
                return importer.GenerateTemplate(fileName);
            }

            {
                using var importer = new Import<T>(_importHeaderFilters, _importResultFilters);
                return importer.GenerateTemplate(fileName);
            }
        }

        public Task<byte[]> GenerateTemplateBytes<T>() where T : class, new()
        {
            var isMultipleSheetType = false;
            var tableType = typeof(T);
            List<PropertyInfo> sheetPropertyList = new List<PropertyInfo>();
            var sheetProperties = tableType.GetProperties();

            foreach (var sheetProperty in sheetProperties)
            {
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    continue;
                }

                if (string.IsNullOrEmpty(importerAttribute.SheetName))
                {
                    continue;
                }

                isMultipleSheetType = true;
                sheetPropertyList.Add(sheetProperty);
            }

            if (isMultipleSheetType)
            {
                using var importer = new ImportMultipleSheet(sheetPropertyList);
                return importer.GenerateTemplateByte();
            }
            else
            {
                using var importer = new Import<T>(_importHeaderFilters, _importResultFilters);
                return importer.GenerateTemplateByte();
            }
        }

        /// <summary>
        ///     导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="labelingFilePath"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(string filePath, string labelingFilePath = null) where T : class, new()
        {
            filePath.CheckExcelFileName();
            using var importer = new Import<T>(_importHeaderFilters, _importResultFilters, filePath, labelingFilePath);
            return importer.ImportExcel();
        }

        /// <summary>
        ///     导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(Stream stream) where T : class, new()
        {
            using var importer = new Import<T>(_importHeaderFilters, _importResultFilters, stream);
            return importer.ImportExcel();
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="labelStream">错误流</param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(Stream stream, Stream labelStream) where T : class, new()
        {
            using var importer = new Import<T>(_importHeaderFilters, _importResultFilters, stream, labelStream);
            return importer.ImportExcel();
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="labelStream"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(string filePath, Stream labelStream) where T : class, new()
        {
            using var importer = new Import<T>(_importHeaderFilters, _importResultFilters, filePath, labelStream);
            return importer.ImportExcel();
        }

        /// <summary>
        /// 导出业务错误数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">文件路径</param>
        /// <param name="bussinessErrorDataList">错误数据</param>
        /// <param name="msg">成功:错误数据返回路径,失败 返回错误原因</param>
        /// <returns></returns>
        public bool OutputBussinessErrorData<T>(string filePath, List<DataRowErrorInfo> bussinessErrorDataList, out string msg) where T : class, new()
        {
            filePath.CheckExcelFileName();
            using var importer = new Import<T>(_importHeaderFilters, _importResultFilters, filePath);
            return importer.OutputBussinessErrorData(bussinessErrorDataList, out msg);
        }

        /// <summary>
        /// 导出业务错误数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">流</param>
        /// <param name="bussinessErrorDataList">错误数据</param>
        /// <param name="fileByte">成功:错误数据返回文件流字节,失败 返回null</param>
        /// <returns></returns>
        public bool OutputBussinessErrorData<T>(Stream stream, List<DataRowErrorInfo> bussinessErrorDataList, out byte[] fileByte) where T : class, new()
        {
            using var importer = new Import<T>(_importHeaderFilters, _importResultFilters);
            return importer.OutputBussinessErrorDataByte(stream, bussinessErrorDataList, out fileByte);
        }

        /// <summary>
        /// 导入多个Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <param name="filePath"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        public async Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(string filePath) where T : class, new()
        {
            filePath.CheckExcelFileName();
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }
            var resultList = new Dictionary<string, ImportResult<object>>();
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();
            using (var importer = new ImportMultipleSheet(filePath))
            {
                for (var i = 0; i < sheetProperties.Length; i++)
                {
                    var sheetProperty = sheetProperties[i];
                    var importerAttribute =
                        (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                    if (importerAttribute == null)
                    {
                        throw new Exception($"{Resource.ExcelimporterAttributeFeaturesIsNotAnnotated}{sheetProperty.Name}");
                    }
                    //if (string.IsNullOrEmpty(importerAttribute.SheetName))
                    //{
                    //    throw new Exception($"Sheet属性{sheetProperty.Name}的ExcelImporterAttribute特性没有设置SheetName");
                    //}
                    bool isSaveLabelingError = i == sheetProperties.Length - 1;
                    //最后一个属性才保存标注的错误,避免多次保存
                    var result = await importer.Import(importerAttribute.SheetName, importerAttribute.SheetIndex, sheetProperty.PropertyType, isSaveLabelingError);
                    resultList.Add(importerAttribute.SheetName ??
                        importerAttribute.SheetIndex.ToString(), result);
                }
            }
            return resultList;
        }

        /// <summary>
        /// 导入多个Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <param name="stream"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        public async Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(Stream stream) where T : class, new()
        {
            var resultList = new Dictionary<string, ImportResult<object>>();
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();
            using var importer = new ImportMultipleSheet(stream);
            for (var i = 0; i < sheetProperties.Length; i++)
            {
                var sheetProperty = sheetProperties[i];
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    throw new Exception($"{Resource.ExcelimporterAttributeFeaturesIsNotAnnotated}{sheetProperty.Name}");
                }
                //if (string.IsNullOrEmpty(importerAttribute.SheetName))
                //{
                //    throw new Exception($"Sheet属性{sheetProperty.Name}的ExcelImporterAttribute特性没有设置SheetName");
                //}
                var isSaveLabelingError = i == sheetProperties.Length - 1;
                //最后一个属性才保存标注的错误,避免多次保存
                var result = await importer.Import(importerAttribute.SheetName, importerAttribute.SheetIndex, sheetProperty.PropertyType, isSaveLabelingError);
                resultList.Add(importerAttribute.SheetName ??
                               importerAttribute.SheetIndex.ToString(), result);
            }

            return resultList;
        }



        /// <summary>
        /// 导入多个相同类型的Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <typeparam name="TSheet">Sheet类</typeparam>
        /// <param name="filePath"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型TSheet</returns>
        public async Task<Dictionary<string, ImportResult<TSheet>>> ImportSameSheets<T, TSheet>(string filePath)
            where T : class, new() where TSheet : class, new()
        {
            filePath.CheckExcelFileName();
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }
            var resultList = new Dictionary<string, ImportResult<TSheet>>();
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();
            using var importer = new ImportMultipleSheet(filePath);
            for (var i = 0; i < sheetProperties.Length; i++)
            {
                var sheetProperty = sheetProperties[i];
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    throw new Exception($"{Resource.ExcelimporterAttributeFeaturesIsNotAnnotated}{sheetProperty.Name}");
                }
                //if (string.IsNullOrEmpty(importerAttribute.SheetName))
                //{
                //    throw new Exception($"Sheet属性{sheetProperty.Name}的ExcelImporterAttribute特性没有设置SheetName");
                //}
                bool isSaveLabelingError = i == sheetProperties.Length - 1;
                //最后一个属性才保存标注的错误,避免多次保存
                var result = await importer.Import(importerAttribute.SheetName, importerAttribute.SheetIndex, sheetProperty.PropertyType, isSaveLabelingError);
                var tResult = new ImportResult<TSheet>();
                tResult.Data = new List<TSheet>();
                if (result.Data.Count > 0)
                {
                    foreach (var item in result.Data)
                    {
                        tResult.Data.Add((TSheet)item);
                    }
                }
                tResult.Exception = result.Exception;
                tResult.RowErrors = result.RowErrors;
                tResult.TemplateErrors = result.TemplateErrors;
                resultList.Add(
                    importerAttribute.SheetName ??
                    importerAttribute.SheetIndex.ToString(),
                    tResult);
            }

            return resultList;
        }

        /// <summary>
        /// 判断Dto类型是否为多Sheet类
        /// </summary>
        /// <typeparam name="T">Dto类型</typeparam>
        /// <returns></returns>
        private bool DtoTypeIsMultipleSheet<T>()
        {
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();

            for (var i = 0; i < sheetProperties.Length; i++)
            {
                var sheetProperty = sheetProperties[i];
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(importerAttribute.SheetName))
                {
                    return true;
                }
            }
            return false;
        }

    }
}