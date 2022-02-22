using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace YunMa.Excel.Utility
{
    /// <summary>
    /// 数据验证帮助类
    /// </summary>
    internal static class ValidatorHelper
    {
        private const string RegexChsChar = "^[\u4e00-\u9fa5_a-zA-Z0-9]+$";
        private const string RegexEngChar = "^[a-zA-Z0-9]+$";
        private const string RegexSplitChar = "^[\u4e00-\u9fa5_a-zA-Z0-9]+(,[\u4e00-\u9fa5_a-zA-Z0-9]+)*$";

        public static bool TryValidate(object obj, out ICollection<ValidationResult> validationResults)
        {
            var context = new ValidationContext(obj, null, null);
            validationResults = new List<ValidationResult>();
            return Validator.TryValidateObject(obj, context, validationResults, true);
        }

        /// <summary>
        /// 校验中文
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static bool ValidChineseInput(this string text)
        {
            var regex = new Regex(RegexChsChar);
            return regex.IsMatch(text);
        }

        /// <summary>
        /// 校验英文
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static bool ValidEnglishInput(this string text)
        {
            var regex = new Regex(RegexEngChar);
            return regex.IsMatch(text);
        }

        public static bool ValidSplitCharInput(this string text)
        {
            var regex = new Regex(RegexSplitChar);
            return regex.IsMatch(text);
        }
    }
}