using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using YunMa.Excel.Core;
using YunMa.Excel.Core.Attributes.Export;
using YunMa.Excel.Core.Extensions;
using YunMa.Excel.Core.Models;
using YunMa.Excel.Exporter.Base.Filter;
using YunMa.Excel.Exporter.Base.Impl;
using YunMa.Excel.Exporter.Base.Utility;

namespace YunMa.Excel.Exporter
{ 
    /// <summary>
    ///     Excel导出程序
    /// </summary>
    public class ExcelExporter : IExcelExporter
    {
        private ExcelPackage _excelPackage;
        private bool _isSeparateColumn;
        private bool _isSeparateBySheet;
        private bool _isSeparateByRow;
        private bool _isAppendHeaders;

        private readonly IEnumerable<IExporterHeaderFilter> _filters;

        public ExcelExporter(IEnumerable<IExporterHeaderFilter> filters)
        {
            _filters = filters;
        }


        /// <summary>
        /// 导出字节
        /// </summary>
        /// <param name="type"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray(DataTable dataItems, Type type)
        {
            var helper = new Export<DataTable>(type, _filters);
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Rows.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var ds = dataItems.SplitDataTable(helper.ExcelExporterSettings.MaxRowNumberOnASheet);
                    var sheetCount = ds.Tables.Count;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = ds.Tables[i];
                        helper.AddExcelWorksheet();
                        helper.ExportExcel(sheetDataItems);
                    }

                    return Task.FromResult(helper.CurrentExcelPackage.GetAsByteArray());
                }
            }

            using var ep = helper.ExportExcel(dataItems);
            return Task.FromResult(ep.GetAsByteArray());
        }

        /// <summary>
        /// 导出字节
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class, new()
        {
            var helper = new Export<T>(_filters);
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var sheetCount = (int)(dataItems.Count / helper.ExcelExporterSettings.MaxRowNumberOnASheet) +
                                     ((dataItems.Count % helper.ExcelExporterSettings.MaxRowNumberOnASheet) > 0
                                         ? 1
                                         : 0);
                    for (var i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = dataItems.Skip(i * helper.ExcelExporterSettings.MaxRowNumberOnASheet)
                            .Take(helper.ExcelExporterSettings.MaxRowNumberOnASheet).ToList();
                        helper.AddExcelWorksheet();
                        helper.ExportExcel(sheetDataItems);
                    }

                    return Task.FromResult(helper.CurrentExcelPackage.GetAsByteArray());
                }
            }
            using var ep = helper.ExportExcel(dataItems);
            return Task.FromResult(ep.GetAsByteArray());
        }

        public Task<byte[]> ExportAsByteArray<T>(DataTable dataItems) where T : class, new()
        {
            var helper = new Export<T>(_filters);
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Rows.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var ds = dataItems.SplitDataTable(helper.ExcelExporterSettings.MaxRowNumberOnASheet);
                    var sheetCount = ds.Tables.Count;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = ds.Tables[i];
                        helper.AddExcelWorksheet();
                        helper.ExportExcel(sheetDataItems);
                    }
                    return Task.FromResult(helper.CurrentExcelPackage.GetAsByteArray());
                }
            }
            else
            {
                using (var ep = helper.ExportExcel(dataItems))
                {
                    return Task.FromResult(ep.GetAsByteArray());
                }
            }
        }

        public Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName = "导出结果")
        {
            var helper = new Export<DataTable>(_filters);
            var headerList = new List<ExporterHeaderInfo>();
            for (var i = 1; i <= items.Length; i++)
            {
                var item = items[i - 1];
                var exporterHeaderInfo =
                    new ExporterHeaderInfo()
                    {
                        Index = i,
                        DisplayName = item,
                        CsTypeName = "string",
                        PropertyName = item,
                        ExporterHeaderAttribute = new ExporterHeaderAttribute(item) { },
                    };
                headerList.Add(exporterHeaderInfo);
            }

            helper.AddExcelWorksheet(sheetName);
            helper.AddExporterHeaderInfoList(headerList);
            using (var ep = helper.ExportHeaders())
            {
                return Task.FromResult(ep.GetAsByteArray());
            }
        }

        public Task<byte[]> ExportAsByteArray(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null,
            int maxRowNumberOnASheet = 1000000)
        {
            var helper = new Export<DataTable>(_filters);
            helper.ExcelExporterSettings.MaxRowNumberOnASheet = maxRowNumberOnASheet;
            helper.SetExporterHeaderFilter(exporterHeaderFilter);

            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Rows.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var ds = dataItems.SplitDataTable(helper.ExcelExporterSettings.MaxRowNumberOnASheet);
                    var sheetCount = ds.Tables.Count;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = ds.Tables[i];
                        helper.AddExcelWorksheet();
                        helper.ExportExcel(sheetDataItems);
                    }
                    return Task.FromResult(helper.CurrentExcelPackage.GetAsByteArray());
                }
            }
            else
            {
                using (var ep = helper.ExportExcel(dataItems))
                {
                    return Task.FromResult(ep.GetAsByteArray());
                }
            }
        }
        public Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class, new()
        {
            var helper = new Export<DataTable>(_filters);
            using var ep = helper.ExportHeaders();
            return Task.FromResult(ep.GetAsByteArray());
        }

        public Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class
        {
            using var helper = new TemplateExport<T>();
            using var sr = new MemoryStream();
            helper.Export(template, data, (package) => { package.SaveAs(sr); });
            return Task.FromResult(sr.ToArray());
        }


        public Task<byte[]> ExportAppendDataAsByteArray()
        {
            if (this._excelPackage == null)
            {
                throw new ArgumentNullException(Resource.AppendMethodMustBeBeforeCurrentMethod);
            }
            var bytes = _excelPackage.GetAsByteArray();
            Reset();
            return Task.FromResult(bytes);
        }


        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="dataItems">数据列</param>
        /// <returns>文件</returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class, new()
        {
            var bytes = await ExportAsByteArray(dataItems);
            return bytes.ToExcelExportFileInfo(fileName);
        }



        /// <summary>
        /// 导出DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, DataTable dataItems) where T : class, new()
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAsByteArray<T>(dataItems);
            return bytes.ToExcelExportFileInfo(fileName);
        }

    
        public Task<ExportFileInfo> ExportByTemplate<T>(string fileName, T data, string template) where T : class
        {
            using (var helper = new TemplateExport<T>())
            {
                var file = new FileInfo(fileName);

                helper.Export(template, data, (package) => { package.SaveAs(file); });
                return Task.FromResult(new ExportFileInfo(file.Name, file.Extension));
            }
        }

      

        public async Task<ExportFileInfo> Export(string fileName, DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null,
            int maxRowNumberOnASheet = 1000000)
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAsByteArray(dataItems, exporterHeaderFilter, maxRowNumberOnASheet);
            return bytes.ToExcelExportFileInfo(fileName);
        }

      
        public ExcelExporter Append<T>(ICollection<T> dataItems, string sheetName = null) where T : class, new()
        {
            var helper = this._excelPackage == null ? new Export<T>(_filters, sheetName) : new Export<T>(_excelPackage, _filters, sheetName);
            if (_isSeparateColumn || _isSeparateBySheet || _isSeparateByRow)
            {
                var name = helper.ExcelExporterSettings?.Name ?? Resource.ExportResult;

                if (this._excelPackage?.Workbook.Worksheets.Any(x => x.Name == name) ?? false)
                {
                    throw new ArgumentNullException($"{Resource.ASheetWithTheNameAlreadyExists}:{name}");
                }
            }

            this._excelPackage = helper.ExportExcel(dataItems);

            if (_isSeparateColumn)
            {
                //#if NET461
                helper.CopySheet(0,
                    1);
                //#else
                //                helper.CopySheet(0,
                //                      1);
                //#endif

                _isSeparateColumn = false;
            }

            if (_isSeparateByRow)
            {
                //#if NET461
                //                helper.CopyRows(0,
                //                    1, _isAppendHeaders);
                //#else
                helper.CopyRows(0,
                    1, _isAppendHeaders);
                //#endif
            }

            _isSeparateBySheet = false;
            _isSeparateByRow = false;
            _isAppendHeaders = false;
            return this;
        }
        /// <summary>
        ///		分割集合到当前Sheet追加Column
        /// </summary>
        /// <returns></returns>
        public ExcelExporter SeparateByColumn()
        {

            if (_excelPackage == null)
            {
                throw new ArgumentNullException(Resource.AppendMethodMustBeBeforeCurrentMethod);
            }

            _isSeparateColumn = true;
            return this;
        }
        /// <summary>
        ///     分割多出多个sheet
        /// </summary>
        /// <returns></returns>
        public ExcelExporter SeparateBySheet()
        {
            if (_excelPackage == null)
            {
                throw new ArgumentNullException(Resource.AppendMethodMustBeBeforeCurrentMethod);
            }

            _isSeparateBySheet = true;
            return this;
        }
        /// <summary>
        ///     追加rows到当前sheet
        /// </summary>
        /// <returns></returns>
        public ExcelExporter SeparateByRow()
        {
            if (_excelPackage == null)
            {
                throw new ArgumentNullException(Resource.AppendMethodMustBeBeforeCurrentMethod);
            }

            _isSeparateByRow = true;
            return this;
        }
       
        public async Task<ExportFileInfo> ExportAppendData(string fileName)
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAppendDataAsByteArray();
            return bytes.ToExcelExportFileInfo(fileName);
        }


        /// <summary>
        ///     追加表头
        /// </summary>
        /// <returns></returns>
        public ExcelExporter AppendHeaders()
        {
            if (_excelPackage == null)
            {
                throw new ArgumentNullException(Resource.AppendMethodMustBeBeforeCurrentMethod);
            }

            if (!_isSeparateByRow)
            {
                throw new ArgumentNullException(Resource.AppendMethodMustBeBeforeCurrentMethod);
            }

            _isAppendHeaders = true;
            return this;
        }
        private void Reset()
        {
            _excelPackage = null;
            _isSeparateByRow = false;
            _isAppendHeaders = false;
            _isSeparateBySheet = false;
            _isSeparateColumn = false;
        }
    }
}