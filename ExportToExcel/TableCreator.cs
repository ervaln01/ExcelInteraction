namespace ExportToExcel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;
    using ExportToExcel.DataProcessing;
    using ExportToExcel.Models;

    /// <summary>
    /// Класс, осуществляющий создание таблицы.
    /// </summary>
    public static class TableCreator
    {
        /// <summary>
        /// Формирование таблицы из списка данных.
        /// </summary>
        /// <param name="values">Перечисление значений.</param>
        /// <param name="showHeader">Флаг показа заголовка.</param>
        /// <param name="sheetName">Название таблицы.</param>
        /// <returns>Экземпляр таблицы.</returns>
        public static Table<T> ToTable<T>(this IEnumerable<T> values, bool showHeader = true, string sheetName = null) => new Table<T>()
        {
            DataSource = values,
            ShowHeader = showHeader,
            Title = sheetName ?? values.First().GetType().Name
        };

        /// <summary>
        /// Добавление колонки.
        /// </summary>
        /// <param name="table">Таблица.</param>
        /// <param name="expression">Выражение.</param>
        /// <param name="title">Заголовок.</param>
        /// <param name="dataType">Тип данных.</param>
        /// <param name="format">Формат.</param>
        /// <param name="convertEmptyStringToNull">Необходимость конвертирования пустой строки в NULL.</param>
        /// <param name="encodeValue">Флаг показывающий закодировано ли значение.</param>
        /// <param name="nullDisplayText">Дефолтный текст.</param>
        /// <param name="culture">Региональный параметр.</param>
        /// <param name="select">Функция выборки.</param>
        /// <returns>Экземпляр таблицы.</returns>
        public static Table<T> AddColumn<T, TValue>(
            this Table<T> table,
            Expression<Func<T, TValue>> expression = null,
            ColumnFormat format = null,
            Func<T, object> select = null)
        {
            var tableColumn = new Column<T>();

            if (expression != null)
            {
                var emptyMetadata = new Metadata();
                var metadata = emptyMetadata.Create(expression);
                tableColumn.DataType = format?.DataType ?? (Type.GetType(metadata.DataType) ?? typeof(string));
                tableColumn.Format = format?.Format ?? metadata.DisplayFormat;
                tableColumn.Title = format?.Title ?? metadata.DisplayName;
                tableColumn.ConvertEmptyStringToNull = format?.ConvertEmptyStringToNull ?? metadata.ConvertEmptyStringToNull;
                tableColumn.NullDisplayText = format?.NullDisplayText ?? metadata.NullDisplayText;
                tableColumn.EncodeValue = format?.EncodeValue ?? metadata.EncodeValue;
            }
            else
            {
                tableColumn.DataType = format?.DataType ?? typeof(string);
                tableColumn.Format = format?.Format;
                tableColumn.Title = format?.Title;

                if (format?.ConvertEmptyStringToNull != null)
                    tableColumn.ConvertEmptyStringToNull = format.ConvertEmptyStringToNull.Value;

                tableColumn.NullDisplayText = format?.NullDisplayText;

                if (format?.EncodeValue != null)
                    tableColumn.EncodeValue = format.EncodeValue.Value;
            }

            if (format?.Culture != null)
                tableColumn.Culture = format?.Culture;

            var unCorrectSelectFunction = select == null && expression != null;
            tableColumn.SelectFunction = !unCorrectSelectFunction ? select : obj => expression.Compile()(obj);

            table.Columns.Add(tableColumn);

            return table;
        }

        /// <summary>
        /// Создание таблицы по заданному пути.
        /// </summary>
        /// <param name="table">Таблица.</param>
        public static void GenerateTable<T>(this Table<T> table, string path)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            if (path == null)
                throw new ArgumentNullException(nameof(path));

            var generator = new TableGenerator<T>(table);
            generator.Generate(path);
        }

        public static Table<T> Concat<T>(this Table<T> table, Table<T> other)
        {
            if (table.Columns.Count != other.Columns.Count)
                throw new InvalidOperationException();

            table.AdditionalTables = table.AdditionalTables.Concat(new[] { other });

            return table;
        }
    }
}