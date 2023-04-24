using System.Collections.Generic;
using ImportFromExcel.Data;

namespace ImportFromExcel.Parsers
{
    public interface ITableParser
    {
        void Parse<T>(T data, IEnumerable<string> cells) where T : ITableData;
    }
}