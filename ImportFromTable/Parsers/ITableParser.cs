using System.Collections.Generic;
using ImportFromTable.Data;

namespace ImportFromTable.Parsers
{
    public interface ITableParser
    {
        void Parse<T>(T data, IEnumerable<string> cells) where T : ITableData;
    }
}