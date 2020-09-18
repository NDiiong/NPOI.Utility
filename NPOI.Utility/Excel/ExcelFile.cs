using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.Utility.Helpers;
using static NPOI.Utility.Excel.ExcelScheme;
using FileIO = System.IO.File;

namespace NPOI.Utility.Excel
{
    public static class ExcelFile
    {
        public static IEnumerable<T> Load<T>(byte[] contentFile, Action<ExcelScheme<T>> schemeBuilder) where T : class, new()
        {
            if (contentFile == null)
                throw new ArgumentNullException(nameof(contentFile));

            using (var inputStream = new MemoryStream(contentFile))
            {
                return ReadExcel(WorkbookFactory.Create(inputStream), schemeBuilder);
            }
        }

        public static IEnumerable<T> Load<T>(string excelPath, Action<ExcelScheme<T>> schemeBuilder) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(excelPath))
                throw new ArgumentNullException(nameof(excelPath));

            if (!FileIO.Exists(excelPath))
                throw new FileNotFoundException(default, Path.GetFileName(excelPath));

            return ReadExcel(WorkbookFactory.Create(excelPath), schemeBuilder);
        }

        private static int GetColumnIndex(IRow headerRow, Column column)
        {
            if (column.Index >= 0) return column.Index;

            var columnIndex = -1;
            if (!string.IsNullOrWhiteSpace(column.Title))
            {
                var cellOfTitle = headerRow.Cells.Find(f => string.Equals(f.StringCellValue, column.Title, StringComparison.OrdinalIgnoreCase));

                if (cellOfTitle != null)
                {
                    columnIndex = cellOfTitle.ColumnIndex;
                    column.Index = columnIndex;
                }
            }
            else if (column.AutoIndex)
            {
                var propertyName = column.PropertyInfo.Name;
                var cellOfTitle = headerRow.Cells.Find(f => string.Equals(f.StringCellValue, propertyName, StringComparison.OrdinalIgnoreCase));

                if (cellOfTitle != null)
                {
                    columnIndex = cellOfTitle.ColumnIndex;
                    column.Index = columnIndex;
                }
            }

            return columnIndex;
        }

        private static ISheet GetSheetWorkbook<T>(this IWorkbook workbook, ExcelScheme<T> excelScheme) where T : class
        {
            if (excelScheme.SheetIndex > -1)
                return workbook.GetSheetAt(excelScheme.SheetIndex);

            if (!string.IsNullOrWhiteSpace(excelScheme.SheetName))
                return workbook.GetSheet(excelScheme.SheetName);

            throw new SheetNotFoundException("Please set the 'SheetIndex' or 'SheetName' for attributes");
        }

        private static IEnumerable<T> ReadExcel<T>(IWorkbook workbook, Action<ExcelScheme<T>> schemeBuilder) where T : class, new()
        {
            if (schemeBuilder == null)
                throw new ArgumentNullException(nameof(schemeBuilder));

            var excelScheme = new ExcelScheme<T>();
            schemeBuilder?.Invoke(excelScheme);

            var sheet = workbook.GetSheetWorkbook(excelScheme);

            var rows = sheet.GetRowEnumerator();

            var headerRow = sheet.GetRow(0);

            var entities = Activator.CreateInstance<List<T>>();

            while (rows.MoveNext())
            {
                var row = rows.Current as IRow;
                if (row.RowNum < excelScheme.StartRow) continue;

                var entity = Activator.CreateInstance<T>();

                foreach (var column in excelScheme.Columns)
                {
                    var columnIndex = GetColumnIndex(headerRow, column);

                    if (columnIndex < 0)
                        throw new CellNotFoundException("Please set the 'index' for attributes");

                    var value = row.GetCellValue(columnIndex, column.PropertyInfo.PropertyType, column.Format);

                    column.PropertyInfo.SetValue(entity, value, default);
                }
                entities.Add(entity);
            }
            return entities;
        }

        private static object GetCellValue(this IRow row, int index, Type propertyType, string format)
        {
            var cell = row.GetCell(index);

            if (cell == null) return null;

            var cellValue = CellValue(cell, format);

            return TypeHelper.ConvertTypeCode(propertyType, cellValue, format);
        }

        private static string CellValue(ICell cell, string format, IFormatProvider provider = default)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return DateTime.FromOADate(cell.NumericCellValue).ToString(format, provider);
                    return cell.NumericCellValue.ToString(provider);

                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();

                default:
                    return string.Empty;
            }
        }
    }
}