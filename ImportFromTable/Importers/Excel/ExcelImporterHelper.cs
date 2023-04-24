using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportFromTable.Data;
using ImportFromTable.Parsers;

namespace ImportFromTable.Importers.Excel
{
    public static class ExcelImporterHelper
    {
        public static IEnumerable<T> GetRowsData<T>(this ITableParser parser, T data, IEnumerable<Row> rows, SharedStringTable sharedStringTable) where T : ITableData
        {
            return rows.Select(row =>
            {
                var clone = (T)data.Clone();
                return parser.ParseRow(clone, row, sharedStringTable);
            });
        }

        public static T ParseRow<T>(this ITableParser parser, T data, Row row, SharedStringTable sharedStringTable) where T : ITableData
        {
            var cells = row.Elements<Cell>();
            var parsedCells = cells.Select(cell => ParseCell(cell, sharedStringTable)).ToArray();
            parser.Parse(data, parsedCells);

            return data;
        }

        public static string ParseCell(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.DataType == null || cell.DataType != CellValues.SharedString)
                return cell.InnerText;

            if (!int.TryParse(cell.InnerText, out var result))
                return string.Empty;

            return sharedStringTable.ElementAt(result).InnerText;
        }
    }
}