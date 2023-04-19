using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportFromExcel.Data;
using ImportFromExcel.Parsers;

namespace ImportFromExcel
{
    public static class TableDisassembler
    {
        private static int _parsedCount = 0;

        public static TableData<T> Import<T>(this IExcelParser parser, T data, string path, bool hasHeaders = true) where T : IExcelData
        {
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                var workbook = document.WorkbookPart;
                var sharedStringTable = workbook?.SharedStringTablePart?.SharedStringTable;

                if (workbook.WorksheetParts.Count() != 1)
                    throw new InvalidOperationException("Невозможно получить единственную таблицу из книги!");

                var worksheet = workbook.WorksheetParts.First().Worksheet;

                var rowDataList = new List<T>();
                _parsedCount = 0;

                foreach (var sheetData in worksheet.Elements<SheetData>())
                {
                    if (!sheetData.HasChildren)
                        continue;

                    var rows = sheetData.Elements<Row>();

                    if (hasHeaders)
                        rows = rows.Skip(1);

                    var rowsData = parser.GetRowsData(data, rows, sharedStringTable);
                    rowDataList.AddRange(rowsData);
                }

                var result = new TableData<T>(rowDataList, _parsedCount);
                _parsedCount = 0;

                return result;
            }
        }

        private static IEnumerable<T> GetRowsData<T>(this IExcelParser parser, T data, IEnumerable<Row> rows, SharedStringTable sharedStringTable) where T : IExcelData
        {
            return rows.Select(row =>
            {
                var clone = (T)data.Clone();
                return parser.ParseRow(clone, row, sharedStringTable);
            });
        }

        private static T ParseRow<T>(this IExcelParser parser, T data, Row row, SharedStringTable sharedStringTable) where T : IExcelData
        {
            var cells = row.Elements<Cell>();
            var parsedCells = cells.Select(cell => ParseCell(cell, sharedStringTable)).ToArray();
            parser.Parse(data, parsedCells);

            if (data.IsParsed)
                _parsedCount++;

            return data;
        }

        private static string ParseCell(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.DataType == null || cell.DataType != CellValues.SharedString)
                return cell.InnerText;

            if (!int.TryParse(cell.InnerText, out var result))
                return string.Empty;

            return sharedStringTable.ElementAt(result).InnerText;
        }
    }
}