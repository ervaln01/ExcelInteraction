using System.Collections.Generic;
using ImportFromExcel.Data;

namespace ImportFromExcel.Parsers
{
    public class SimpleParser : IExcelParser
    {
        public void Parse<T>(T data, IEnumerable<string> cells) where T : IExcelData => data.Parse(cells);
    }
}