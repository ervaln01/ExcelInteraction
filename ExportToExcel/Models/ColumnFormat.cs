using System;
using System.Globalization;

namespace ExportToExcel.Models
{
    public class ColumnFormat
    {
        public string Title { get; private set; }
        public Type DataType { get; private set; }
        public string Format { get; private set; }
        public bool? ConvertEmptyStringToNull { get; private set; }
        public bool? EncodeValue { get; private set; }
        public string NullDisplayText { get; private set; }
        public CultureInfo Culture { get; private set; }

        public ColumnFormat()
        {
            Title = null;
            DataType = null;
            Format = null;
            ConvertEmptyStringToNull = null;
            EncodeValue = null;
            NullDisplayText = null;
            Culture = null;
        }

        public ColumnFormat SetTitle(string title)
        {
            Title = title;
            return this;
        }

        public ColumnFormat SetDataType(Type dataType)
        {
            DataType = dataType;
            return this;
        }

        public ColumnFormat SetFormat(string format)
        {
            Format = format;
            return this;
        }

        public ColumnFormat SetConvertEmptyStringToNull(bool? convertEmptyStringToNull)
        {
            ConvertEmptyStringToNull = convertEmptyStringToNull;
            return this;
        }

        public ColumnFormat SetEncodeValue(bool? encodeValue)
        {
            EncodeValue = encodeValue;
            return this;
        }

        public ColumnFormat SetNullDisplayText(string nullDisplayText)
        {
            NullDisplayText = nullDisplayText;
            return this;
        }

        public ColumnFormat SetCulture(CultureInfo culture)
        {
            Culture = culture;
            return this;
        }
    }

}
