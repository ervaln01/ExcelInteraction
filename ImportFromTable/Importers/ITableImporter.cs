using ImportFromTable.Data;
using ImportFromTable.Parsers;

namespace ImportFromTable.Importers
{
    public interface ITableImporter
    {
        ParsedTableInfo<T> Import<T>(ITableParser parser, T data, string path, bool hasHeaders = true) where T : ITableData;
    }
}