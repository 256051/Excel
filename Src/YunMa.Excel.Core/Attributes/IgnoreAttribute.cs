﻿using System;

namespace YunMa.Excel.Core.Attributes
{
    /// <summary>
    ///     忽略特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class IgnoreAttribute : Attribute
    {
        /// <summary>
        ///     是否忽略导入，默认true
        /// </summary>
        public bool IsImportIgnore { get; set; } = true;

        /// <summary>
        ///     是否忽略导出，默认true
        /// </summary>
        public bool IsExportIgnore { get; set; } = true;
    }
}