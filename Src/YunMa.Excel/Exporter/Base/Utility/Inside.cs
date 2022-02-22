namespace YunMa.Excel.Exporter.Base.Utility
{
    internal class Inside<T> where T : class
    {
        private bool? _isDictionaryType;

        private bool? _isExpandoObjectType;


        private bool? _isJObjectType;

        /// <summary>
        ///     是否是支持的动态类型（JObject、Dictionary（仅支持key为string类型））
        /// </summary>
        /// <remarks>支持DataTable</remarks>
        public bool IsDynamicSupportTypes => IsDictionaryType || IsJObjectType || IsExpandoObjectType;
        /// <summary>
        ///     是否是符合要求的字典类型
        /// </summary>
        public bool IsDictionaryType
        {
            get
            {
                if (_isDictionaryType.HasValue) return _isDictionaryType.Value;

                var name = typeof(T).Name;
                _isDictionaryType = name switch
                {
                    "Dictionary`2" => typeof(T).GetGenericArguments()[0] == typeof(string),
                    _ => false
                };
                return _isDictionaryType.Value;
            }
        }

        public bool IsExpandoObjectType
        {
            get
            {
                if (_isExpandoObjectType.HasValue) return _isExpandoObjectType.Value;
                _isExpandoObjectType = typeof(T).Name == "ExpandoObject";
                return _isExpandoObjectType.Value;
            }
        }


        /// <summary>
        ///     是否是JObject类型
        /// </summary>
        public bool IsJObjectType
        {
            get
            {
                if (_isJObjectType.HasValue) return _isJObjectType.Value;

                var name = typeof(T).Name;
                _isJObjectType = name switch
                {
                    "JObject" => true,
                    _ => false
                };
                return _isJObjectType.Value;
            }
        }
    }
}