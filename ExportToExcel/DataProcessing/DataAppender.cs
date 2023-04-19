namespace ExportToExcel.DataProcessing
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using ExportToExcel.Helpers;
    using ExportToExcel.Models;

    /// <summary>
    /// Класс, осуществляющий заполнение таблицы данными.
    /// </summary>
    public class DataAppender<T>
    {
        /// <summary>
        /// Алфавит названий колонок таблицы.
        /// </summary>
        private const string COLUMN_NAMES = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        /// <summary>
        /// Размер алфавита названий колонок таблицы.
        /// </summary>
        private const int COLUMN_NAMES_COUNT = 26;

        /// <summary>
        /// Таблица.
        /// </summary>
        private readonly Table<T> _table;

        /// <summary>
        /// Словарь стилей.
        /// </summary>
        private readonly Dictionary<Column<T>, uint> _styleIndex;

        /// <summary>
        /// Конструктор класса <see cref="DataAppender"/>.
        /// </summary>
        /// <param name="table">Таблица.</param>
        /// <param name="styleIndex">Словарь стилей.</param>
        public DataAppender(Table<T> table, Dictionary<Column<T>, uint> styleIndex)
        {
            _table = table;
            _styleIndex = styleIndex;
        }

        /// <summary>
        /// Добавление заголовков в первую строку таблицы.
        /// </summary>
        /// <param name="sheetData">Данные.</param>
        /// <param name="rowIndex">Индекс строки.</param>
        /// <returns>Кортеж.</returns>
        public Row AppendHeaders(SheetData sheetData, uint rowIndex)
        {
            var row = new Row { RowIndex = new UInt32Value(rowIndex + 1) };

            _table.Columns.ForEach((column, index) =>
            {
                var cell = new Cell
                {
                    CellReference = GetColumnIndex((uint)index) + row.RowIndex,
                    DataType = CellValues.InlineString
                };

                var child = CreateInlineString(string.Format(CultureInfo.InvariantCulture, "{0}", column.Title));
                cell.AppendChild(child);
                cell.StyleIndex = 2u;
                row.AppendChild(cell);
            });

            sheetData.AppendChild(row);

            return row;
        }

        /// <summary>
        /// Добавление данных в строку.
        /// </summary>
        /// <param name="sheetData">Данные.</param>
        /// <param name="rowIndex">Индекс строки.</param>
        /// <param name="value">Значение.</param>
        /// <returns>Кортеж.</returns>
        public Row AppendRow(SheetData sheetData, uint rowIndex, T value)
        {
            var row = new Row { RowIndex = new UInt32Value(rowIndex + 1) };

            _table.Columns.ForEach((column, index) => AppendCell(row, row.RowIndex, (uint)index, column, value));
            sheetData.AppendChild(row);

            return row;
        }

        /// <summary>
        /// Заполнение ячейки.
        /// </summary>
        /// <param name="row">Строка</param>
        /// <param name="rowIndex">Индекс строки.</param>
        /// <param name="columnIndex">Индекс колонки.</param>
        /// <param name="column">Колонка.</param>
        /// <param name="value">Значение.</param>
        /// <returns>Ячейка.</returns>
        private void AppendCell(Row row, uint rowIndex, uint columnIndex, Column<T> column, T value)
        {
            if (value == null)
                return;

            var displayValue = column.GetValue(value);

            if (displayValue == null)
                return;

            var cell = new Cell() { CellReference = GetColumnIndex(columnIndex) + rowIndex };

            var datatype = GetDataType(column.DataType ?? displayValue.GetType());

            switch (datatype)
            {
                case CellValues.Boolean:
                    var boolean = Convert.ToBoolean(displayValue, column.Culture);
                    cell.DataType = CellValues.Boolean;
                    cell.CellValue = new CellValue(Convert.ToInt32(boolean));
                    break;

                case CellValues.Number:
                    var number = Convert.ToString(displayValue, column.Culture);
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(number);
                    break;

                case CellValues.String:
                    var formula = Convert.ToString(displayValue, column.Culture);
                    cell.DataType = CellValues.InlineString;
                    cell.CellFormula = new CellFormula(formula);
                    break;

                case CellValues.InlineString:
                    cell.DataType = CellValues.InlineString;
                    var chlid = CreateInlineString(Convert.ToString(displayValue, column.Culture));
                    cell.AppendChild(chlid);
                    break;

                case CellValues.Date:
                    var dateTime = Convert.ToDateTime(displayValue, column.Culture);
                    cell.CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                    cell.StyleIndex = 1;
                    break;

                default:
                    throw new FormatException("Формат значения ячейки не определён");
            }

            if (column.Format != null)
                cell.StyleIndex = _styleIndex[column];

            row.AppendChild(cell);
        }

        /// <summary>
        /// Получение символьного индекса по числовому значению.
        /// </summary>
        /// <param name="columnIndex">Индекс колонки.</param>
        /// <returns>Символьный индекс.</returns>
        private static string GetColumnIndex(uint columnIndex)
        {
            var sb = new StringBuilder();

            while (columnIndex != 0)
            {
                var index = (int)(columnIndex % COLUMN_NAMES_COUNT);
                sb.Append(COLUMN_NAMES[index]);
                columnIndex /= COLUMN_NAMES_COUNT;
            }

            return sb.ToString();
        }

        /// <summary>
        /// Получение типа данных.
        /// </summary>
        /// <param name="type">Тип.</param>
        /// <returns>Значения ячеек.</returns>
        private static CellValues GetDataType(Type type)
        {
            if (type == null)
                throw new ArgumentNullException(nameof(type));

            var typeCode = Type.GetTypeCode(type);

            switch (typeCode)
            {
                case TypeCode.Empty:
                case TypeCode.DBNull:
                case TypeCode.Object:
                    return CellValues.String;

                case TypeCode.Boolean:
                    return CellValues.Boolean;

                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return CellValues.Number;

                case TypeCode.DateTime:
                    return CellValues.Date;

                case TypeCode.Char:
                case TypeCode.String:
                    return CellValues.InlineString;

                default:
                    throw new FormatException("Формат не определён");
            }
        }

        private InlineString CreateInlineString(string value)
        {
            var text = new Text(value);
            return new InlineString(text);
        }
    }
}