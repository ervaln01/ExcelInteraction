using System.Collections.Generic;
using ImportFromExcel.Data;

namespace ImportFromExcel
{
    public class TableData<T> where T : IExcelData
    {
        public List<T> Data { get; private set; }

        public int Count => Data.Count;

        public int ParsedCount { get; private set; }

        public T this[int index] => Data[index];

        public TableData(List<T> data, int parsedCount)
        {
            Data = data;
            ParsedCount = parsedCount;
        }
    }
}