using ExportToExcel;
using System;
using System.Collections.Generic;

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
			var exampleList = new List<ExampleClass>()
			{
				new() {Count = 1, Name = "John", Date = DateTime.Now.AddDays(-1)},
				new() {Count = 10, Name = "Alex", Date = DateTime.Now},
				new() {Count = 100, Name = "Morgan", Date = DateTime.Now.AddDays(1)},
			};
			exampleList.ToTable(true, "Example data")
				.AddColumn(x => x.Name, "Human name")
				.AddColumn(x => x.Count)
				.AddColumn(x => x.Date, "Date", typeof(DateTime), "dd.MM.yyyy")
				.GenerateTable("example", $"Example table.xlsx");
		}
	}

	public class ExampleClass
	{
		public int Count;
		public string Name;
		public DateTime Date;
	}
}