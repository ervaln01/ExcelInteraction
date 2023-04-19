namespace ExportToExcel.DataProcessing
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using ExportToExcel.Helpers;
    using ExportToExcel.Models;

    /// <summary>
    /// Класс, осуществляющий генерацию таблицы.
    /// </summary>
    public class TableGenerator<T>
    {
        /// <summary>
        /// Таблица.
        /// </summary>
        private readonly Table<T> _table;

        private uint _rowIndex;

        /// <summary>
        /// Словарь стилей.
        /// </summary>
        private readonly Dictionary<Column<T>, uint> _styleIndex;

        /// <summary>
        /// Конструктор класса <see cref="SpreadsheetGenerator{T}"/>
        /// </summary>
        /// <param name="table">Таблица.</param>
        public TableGenerator(Table<T> table)
        {
            _table = table;
            _styleIndex = new Dictionary<Column<T>, uint>();
        }

        /// <summary>
        /// Генерирует таблицу по документу из заданного пути.
        /// </summary>
        /// <param name="path">Путь.</param>
        public void Generate(string path)
        {
            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
                Generate(document);
        }

        /// <summary>
        /// Генерирует таблицу из заданного документа.
        /// </summary>
        /// <param name="document">Документ хранящий таблицу.</param>
        private void Generate(SpreadsheetDocument document)
        {
            if (_table.Columns.Count == 0)
                return;

            _rowIndex = 0;
            var part = document.AddWorkbookPart();
            part.Workbook = new Workbook();

            CreateStyles(part, _table);

            var worksheetPart = part.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();

            CreateSheets(document, worksheetPart, _table);

            var columns = CreateColumns(_table);
            worksheetPart.Worksheet.Append(columns);

            var sheetData = new SheetData();

            sheetData = CreateSheetData(sheetData, _table);

            if (_table.AdditionalTables != null)
                _table.AdditionalTables.ForEach(t => CreateSheetData(sheetData, t));

            worksheetPart.Worksheet.AppendChild(sheetData);
            part.Workbook.Save();
        }

        private void CreateStyles(WorkbookPart part, Table<T> table)
        {
            var stylesCreator = new StylesCreator<T>(table, _styleIndex);
            var stylePart = part.AddNewPart<WorkbookStylesPart>();
            stylePart.Stylesheet = stylesCreator.CreateStylesheet();
        }

        private void CreateSheets(SpreadsheetDocument document, WorksheetPart part, Table<T> table)
        {
            var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet
            {
                Id = document.WorkbookPart.GetIdOfPart(part),
                SheetId = 1,
                Name = table.Title
            };

            sheets.Append(sheet);
        }

        private SheetData CreateSheetData(SheetData sheetData, Table<T> table)
        {
            var appender = new DataAppender<T>(table, _styleIndex);

            if (table.ShowHeader)
                appender.AppendHeaders(sheetData, _rowIndex++);

            if (table.DataSource != null)
                table.DataSource.ForEach(item => appender.AppendRow(sheetData, _rowIndex++, item));

            return sheetData;
        }

        private Columns CreateColumns(Table<T> table) => new Columns(CreateColumn(table));

        private Column CreateColumn(Table<T> table) => new Column
        {
            Min = 1,
            Max = UInt32Value.FromUInt32((uint)table.Columns.Count),
            Width = 20,
            CustomWidth = true
        };
    }
}