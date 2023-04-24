using System.Collections.Generic;
using ImportFromTable.Data;

namespace ImportFromTable
{
    public class ParsedTableInfo<T> where T : ITableData
    {
        public List<T> Data { get; private set; }

        public int Count => Data.Count;

        public int ParsedCount { get; private set; }

        public T this[int index] => Data[index];

        public ParsedTableInfo(List<T> data, int parsedCount)
        {
            Data = data;
            ParsedCount = parsedCount;
        }
    }
}