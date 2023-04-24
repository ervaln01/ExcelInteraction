using System;
using System.Collections.Generic;
using System.Linq;
using ImportFromTable.Data;

namespace ImportFromTable.Importers.Csv
{
    public static class CsvImporterHelper
    {
        private static readonly char[] _defaultDelimeters = { '\t', ';', ',', '.', ' ' };

        public static char GetDelimeter<T>(this T data, string firstLine, string lastLine) where T : ITableData
        {
            var delimeter = data.GetDelimeter(firstLine, lastLine, _defaultDelimeters);

            if (delimeter.HasValue)
                return delimeter.Value;

            var delimeterGroups = firstLine.GroupBy(x => x).Where(g => g.Count() > 1).Select(g => g.Key);
            delimeter = data.GetDelimeter(firstLine, lastLine, delimeterGroups);

            if (delimeter.HasValue)
                return delimeter.Value;

            throw new InvalidOperationException($"Не удалось выбрать подходящий разделитель{Environment.NewLine}Выбран некорректный тип или парсер");
        }

        private static char? GetDelimeter<T>(this T data, string firstLine, string lastLine, IEnumerable<char> delimeters) where T : ITableData
        {
            foreach (var delimeter in delimeters)
            {
                try
                {
                    if (!data.LineParsed(firstLine, delimeter))
                        continue;

                    if (!data.LineParsed(lastLine, delimeter))
                        continue;

                    return delimeter;
                }
                catch
                {
                    continue;
                }
            }

            return null;
        }

        private static bool LineParsed<T>(this T data, string line, char delimeter) where T : ITableData
        {
            var clone = (T)data.Clone();
            clone.Parse(line.Split(delimeter));

            return clone.IsParsed;
        }
    }
}