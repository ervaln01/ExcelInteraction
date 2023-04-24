using System.Collections.Generic;
using System.IO;
using System.Linq;
using ImportFromExcel.Data;
using ImportFromExcel.Parsers;

namespace ImportFromExcel.Importers.Csv
{
    public class CsvImporter : ITableImporter
    {
        public ParsedTableInfo<T> Import<T>(ITableParser parser, T data, string path, bool hasHeaders = true) where T : ITableData
        {
            var lines = File.ReadAllLines(path);
            var delimeter = data.GetDelimeter(lines.First(), lines.Last());

            var rowDataList = new List<T>();

            foreach(var line in lines.Skip(hasHeaders ? 1 : 0))
            {
                var clone = (T)data.Clone();
                clone.Parse(line.Split(delimeter));
                rowDataList.Add(clone);
            }

            var result = new ParsedTableInfo<T>(rowDataList, rowDataList.Count(r => r.IsParsed));

            return result;
        }
    }
}