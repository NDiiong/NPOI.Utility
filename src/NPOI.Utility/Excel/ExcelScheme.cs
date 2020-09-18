using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace NPOI.Utility.Excel
{
    public abstract class ExcelScheme
    {
        internal static readonly Dictionary<string, List<Column>> _columnsCache = new Dictionary<string, List<Column>>(StringComparer.Ordinal);

        internal class Column
        {
            public bool AutoIndex { get; set; }

            public string Format { get; set; }

            public int Index { get; set; }

            public PropertyInfo PropertyInfo { get; set; }

            public string Title { get; set; }
        }
    }

    public class ExcelScheme<T> : ExcelScheme where T : class
    {
        private readonly object _balanceLock = new object();

        public ExcelScheme() => InitDefaultColumns();

        public string Extension { get; set; }

        public int SheetIndex { get; set; } = -1;

        public string SheetName { get; set; }

        public int StartRow { get; set; }

        internal IEnumerable<Column> Columns => _columnsCache[typeof(T).FullName];

        private void InitDefaultColumns()
        {
            if (!_columnsCache.ContainsKey(typeof(T).FullName))
            {
                var columns = new List<Column>();
                foreach (var property in typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty))
                {
                    var attributeSheetColumn = property.GetCustomAttributes(typeof(SheetColumnAttribute), inherit: true).FirstOrDefault();

                    if (attributeSheetColumn is SheetColumnAttribute _sheetColumnTitleAttribute && !_sheetColumnTitleAttribute.IsIgnored)
                    {
                        columns.Add(new Column
                        {
                            PropertyInfo = property,
                            AutoIndex = _sheetColumnTitleAttribute.AutoIndex,
                            Index = _sheetColumnTitleAttribute.Index,
                            Title = _sheetColumnTitleAttribute.Title,
                            Format = _sheetColumnTitleAttribute.Formatter,
                        });
                    }
                }
                lock (_balanceLock)
                    _columnsCache.Add(typeof(T).FullName, columns);
            }
        }
    }
}