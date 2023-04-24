using System;
using System.Collections.Generic;

namespace ImportFromExcel.Data
{
    public interface ITableData : ICloneable
    {
        bool IsParsed { get; set; }
        void Parse(IEnumerable<string> cells);
    }
}