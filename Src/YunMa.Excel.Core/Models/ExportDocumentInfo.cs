using System;
using System.Collections.Generic;
using System.Data;
using YunMa.Excel.Core.Attributes.Export;
using YunMa.Excel.Core.Extensions;

namespace YunMa.Excel.Core.Models
{
    /// <summary>
    /// </summary>
    /// <typeparam name="TData"></typeparam>
    public class ExportDocumentInfoOfListData<TData> where TData : class
    {
        /// <summary>
        /// </summary>
        public ExportDocumentInfoOfListData(ICollection<TData> datas)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Datas = datas;
            Title = typeof(TData).GetAttribute<ExporterAttribute>()?.Name ?? typeof(TData).Name;

            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exporterHeader = propertyInfo.GetAttribute<ExporterHeaderAttribute>() ??
                                     new ExporterHeaderAttribute
                                     {
                                         DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                     };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        ///     文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        public ICollection<TData> Datas { get; set; }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        public DataTable ToDataTable()
        {
            return Datas.ToDataTable();
        }
    }
    /// <summary>
    /// </summary>
    public class ExportDocumentInfo
    {
        /// <summary>
        /// </summary>
        public ExportDocumentInfo(object data, Type type)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Data = data;
            Title = type.GetAttribute<ExporterAttribute>()?.Name ?? type.Name;

            foreach (var propertyInfo in type.GetProperties())
            {
                var exporterHeader = propertyInfo.PropertyType.GetAttribute<ExporterHeaderAttribute>() ??
                                     new ExporterHeaderAttribute
                                     {
                                         DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                     };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        ///     文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        public object Data { get; set; }
    }

    /// <summary>
    /// </summary>
    /// <typeparam name="TData"></typeparam>
    public class ExportDocumentInfo<TData> where TData : class
    {
        /// <summary>
        /// </summary>
        public ExportDocumentInfo(TData data)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Data = data;
            Title = typeof(TData).GetAttribute<ExporterAttribute>()?.Name ?? typeof(TData).Name;

            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exporterHeader = propertyInfo.PropertyType.GetAttribute<ExporterHeaderAttribute>() ??
                                     new ExporterHeaderAttribute
                                     {
                                         DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                     };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        ///     文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        public TData Data { get; set; }
    }
}