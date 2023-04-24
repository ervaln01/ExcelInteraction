using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportFromExcel.Data;
using ImportFromExcel.Parsers;

namespace ImportFromExcel.Importers.Excel
{
    public class ExcelImporter : ITableImporter
    {
        public ParsedTableInfo<T> Import<T>(ITableParser parser, T data, string path, bool hasHeaders = true) where T : ITableData
        {
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                var workbook = document.WorkbookPart;
                var sharedStringTable = workbook.SharedStringTablePart.SharedStringTable;

                if (workbook.WorksheetParts.Count() != 1)
                    throw new InvalidOperationException("Невозможно получить единственную таблицу из книги!");

                var worksheet = workbook.WorksheetParts.First().Worksheet;

                var rowDataList = new List<T>();

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

                var result = new ParsedTableInfo<T>(rowDataList, rowDataList.Count(r => r.IsParsed));

                return result;
            }
        }
    }
}