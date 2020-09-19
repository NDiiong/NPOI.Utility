using System;

namespace NPOI.Utility.Helpers
{
    internal static class TypeHelper
    {
        internal static object ConvertTypeCode(Type propertyType, string @value, string format, IFormatProvider formatProvider)
        {
            if (propertyType == typeof(int))
                return int.Parse(@value, formatProvider);

            if (propertyType == typeof(int?))
                return !string.IsNullOrWhiteSpace(@value) ? int.Parse(@value, formatProvider) : default(int?);

            if (propertyType == typeof(long))
                return long.Parse(@value, formatProvider);

            if (propertyType == typeof(long?))
                return !string.IsNullOrWhiteSpace(@value) ? long.Parse(@value, formatProvider) : default(long?);

            if (propertyType == typeof(bool))
                return bool.Parse(@value);

            if (propertyType == typeof(bool?))
                return !string.IsNullOrWhiteSpace(@value) ? bool.Parse(@value) : default(bool?);

            if (propertyType == typeof(decimal?))
                return !string.IsNullOrWhiteSpace(@value) ? decimal.Parse(@value, formatProvider) : default(decimal?);

            if (propertyType == typeof(decimal))
                return decimal.Parse(@value, formatProvider);

            if (propertyType == typeof(double?))
                return !string.IsNullOrWhiteSpace(@value) ? double.Parse(@value, formatProvider) : default(double?);

            if (propertyType == typeof(double))
                return double.Parse(@value, formatProvider);

            if (propertyType == typeof(DateTime))
                return !string.IsNullOrWhiteSpace(@value) ? DateTime.ParseExact(@value, format, formatProvider) : default;

            if (propertyType == typeof(DateTime?))
            {
                return !string.IsNullOrWhiteSpace(@value)
                    ? DateTime.ParseExact(@value, format, formatProvider)
                    : default(DateTime?);
            }

            return @value;
        }
    }
}