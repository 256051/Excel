
using System;

namespace YunMa.Excel.Core.Attributes
{
    /// <summary>
    ///     值映射
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ValueMappingAttribute : Attribute
    {
        /// <summary>
        ///     设置文本和值映射
        /// </summary>
        /// <param name="text">文本</param>
        /// <param name="value">值</param>
        public ValueMappingAttribute(string text, object value)
        {
            Text = text;
            Value = value;
        }

        /// <summary>
        ///     文本
        /// </summary>
        public string Text { get; }

        /// <summary>
        ///     值
        /// </summary>
        public object Value { get; }
    }
}