using System.Collections.Generic;
using ImportFromTable.Data;

namespace ImportFromTable.Parsers
{
    public class SimpleParser : ITableParser
    {
        public void Parse<T>(T data, IEnumerable<string> cells) where T : ITableData => data.Parse(cells);
    }
}