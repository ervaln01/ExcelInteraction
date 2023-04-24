using System.Collections.Generic;
using ImportFromExcel.Data;

namespace ImportFromExcel.Parsers
{
    public class SimpleParser : ITableParser
    {
        public void Parse<T>(T data, IEnumerable<string> cells) where T : ITableData => data.Parse(cells);
    }
}