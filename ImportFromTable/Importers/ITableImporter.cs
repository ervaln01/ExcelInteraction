using ImportFromExcel.Data;
using ImportFromExcel.Parsers;

namespace ImportFromExcel.Importers
{
    public interface ITableImporter
    {
        ParsedTableInfo<T> Import<T>(ITableParser parser, T data, string path, bool hasHeaders = true) where T : ITableData;
    }
}