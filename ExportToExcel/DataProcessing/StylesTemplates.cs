namespace ExportToExcel.DataProcessing
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class StyleTemplates
    {
        public static Border GetEmptyBorder() => new Border
        {
            LeftBorder = new LeftBorder(),
            RightBorder = new RightBorder(),
            TopBorder = new TopBorder(),
            BottomBorder = new BottomBorder(),
            DiagonalBorder = new DiagonalBorder()
        };

        public static CellStyle GetEmptyCellStyle() => new CellStyle
        {
            Name = "Normal",
            FormatId = 0,
            BuiltinId = 0
        };

        public static CellFormat GetCellFormat(UInt32Value numberFormatId, UInt32Value fontId, bool applyNumberFormat = false, Alignment alignment = null)
        {
            return new CellFormat
            {
                NumberFormatId = numberFormatId,
                FontId = fontId,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = applyNumberFormat,
                Alignment = alignment
            };
        }

        public static Font GetFont(Bold bold = null) => new Font
        {
            Bold = bold,
            FontSize = new FontSize { Val = 11 },
            Color = new Color { Theme = 1 },
            FontName = new FontName { Val = "Calibri" },
            FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
            FontScheme = new FontScheme { Val = FontSchemeValues.Minor }
        };
    }
}