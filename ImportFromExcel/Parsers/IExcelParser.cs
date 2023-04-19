using System.Collections.Generic;
using ImportFromExcel.Data;

namespace ImportFromExcel.Parsers
{
    public interface IExcelParser
    {
        void Parse<T>(T data, IEnumerable<string> cells) where T : IExcelData;
    }
}