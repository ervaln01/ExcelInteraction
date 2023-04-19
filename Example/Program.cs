using ExportToExcel;
using ExportToExcel.Models;
using ImportFromExcel;
using ImportFromExcel.Data;
using ImportFromExcel.Parsers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;


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
			var tableName = "Example table.xlsx";

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
				.GenerateTable(tableName);

			var exampleData = TableDisassembler.Import(new SimpleParser(), new ExampleClass(), tableName, hasHeaders: true);
			exampleData.Data.ForEach(d => Console.WriteLine(d.Show()));

			Console.ReadLine();
		}
	}

	public class ExampleClass : IExcelData
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