using System;

namespace NPOI.Utility.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ColumnAttribute : Attribute
    {
        public bool AutoIndex { get; set; }

        public string Formatter { get; set; }

        public int Index { get; set; } = -1;

        public bool IsIgnored { get; set; }

        public string Title { get; set; }
    }
}