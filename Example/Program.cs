using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ExportToExcel;
using ExportToExcel.Models;
using ImportFromTable.Data;
using ImportFromTable.Importers;
using ImportFromTable.Importers.Csv;
using ImportFromTable.Importers.Excel;
using ImportFromTable.Parsers;

/// <summary>
/// Для работы библиотеки необходимо установить Nuget DocumentFormat.OpenXML
/// Для создания таблицы необходимо использовать 3 метода:
///		--ToTable<T> - создаёт шаблон для таблицы
///		--AddColumn<T> - добавляет в таблицу колонку
///		--GenerateTable<T> - генерирует таблицу
/// </summary>

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = "Example table.xlsx";

            var now = DateTime.Now;
            var exampleList = new List<ExampleClass>()
            {
                new() { Count = 1, Name = "John", Date = now.AddDays(-1) },
                new() { Count = 10, Name = "Alex", Date = now },
                new() { Count = 100, Name = "Morgan", Date = now.AddDays(1) },
            };

            exampleList.ToTable(true, "Example data")
                .AddColumn(x => x.Name, new ColumnFormat().SetTitle("Human name"))
                .AddColumn(x => x.Count)
                .AddColumn(x => x.Date, new ColumnFormat().SetTitle("Date").SetDataType(typeof(string)).SetFormat("dd.MM.yyyy"))
                .GenerateTable(fileName);

            var importer = GetTableImporter(fileName);

            var exampleData = importer.Import(new SimpleParser(), new ExampleClass(), fileName, hasHeaders: true);
            exampleData.Data.ForEach(d => Console.WriteLine(d.Show()));

            Console.ReadLine();
        }

        private static ITableImporter GetTableImporter(string fileName)
        {
            var ext = Path.GetFileNameWithoutExtension(fileName);

            return ext switch
            {
                ".xlsx" => new ExcelImporter(),
                ".csv" or ".txt" => new CsvImporter(),
                _ => throw new InvalidOperationException($"Extension ({ext}) not supported"),
            };
        }
    }

    public class ExampleClass : ITableData
    {
        private const int PARSED_CELLS_COUNT = 3;

        public int Count { get; set; }
        public string Name { get; set; }
        public DateTime Date { get; set; }
        public bool IsParsed { get; set; }

        public object Clone() => new ExampleClass();

        public void Parse(IEnumerable<string> cells)
        {
            try
            {
                if (cells.Count() < PARSED_CELLS_COUNT)
                    throw new InvalidCastException();

                var parsedCells = cells.ToList();

                Count = int.Parse(parsedCells[1]);
                Name = parsedCells[0];
                Date = DateTime.ParseExact(parsedCells[2], "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                IsParsed = true;
            }
            catch
            {
                IsParsed = false;
            }
        }

        public string Show() => $"Count - {Count}; Name - {Name}; Date - {Date}";
    }
}