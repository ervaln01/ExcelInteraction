namespace ExportToExcel.DataProcessing
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using ExportToExcel.Models;

    /// <summary>
    /// Класс задания стилей таблицы.
    /// </summary>
    public class StylesCreator<T>
    {
        private const string MC_URI_DECLARATOIN = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private const string X14AC_URI_DECLARATOIN = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

        /// <summary>
        /// Отступ для установки NumberFormatId.
        /// </summary>
        private const uint NUM_FMT_ID_INDENT = 164;

        /// <summary>
        /// Тапблица.
        /// </summary>
        private readonly Table<T> _table;

        /// <summary>
        /// Словарь стилей.
        /// </summary>
        private readonly Dictionary<Column<T>, uint> _styleIndex;

        /// <summary>
        /// Конструктор класса <see cref="StylesCreator{T}"/>
        /// </summary>
        /// <param name="table">Таблица.</param>
        /// <param name="styleIndex">Словарь стилей.</param>
        public StylesCreator(Table<T> table, Dictionary<Column<T>, uint> styleIndex)
        {
            _table = table;
            _styleIndex = styleIndex;
        }

        /// <summary>
        /// Создание таблицы стилей.
        /// </summary>
        /// <returns>Таблица стилей.</returns>
        public Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet
            {
                MCAttributes = new MarkupCompatibilityAttributes
                {
                    Ignorable = "x14ac"
                }
            };

            stylesheet.AddNamespaceDeclaration("mc", MC_URI_DECLARATOIN);
            stylesheet.AddNamespaceDeclaration("x14ac", X14AC_URI_DECLARATOIN);

            stylesheet.Borders = CreateBorders();
            stylesheet.CellFormats = CreateCellFormats();
            stylesheet.CellStyleFormats = CreateCellStyleFormats();
            stylesheet.CellStyles = CreateCellStyles();
            stylesheet.DifferentialFormats = CreateDifferentialFormats();
            stylesheet.Fills = CreateFills();
            stylesheet.Fonts = CreateFonts();
            stylesheet.NumberingFormats = CreateNumeringFormats();
            stylesheet.StylesheetExtensionList = CreateStylesheetExtensionList();
            stylesheet.TableStyles = CreateTableStyles();

            return stylesheet;
        }

        /// <summary>
        /// Создание рамок.
        /// </summary>
        /// <returns>Рамки.</returns>
        private static Borders CreateBorders()
        {
            var borders = new Borders { Count = 1 };
            borders.Append(StyleTemplates.GetEmptyBorder());

            return borders;
        }

        /// <summary>
        /// Создание форматов ячеек.
        /// </summary>
        /// <returns>Форматы ячеек.</returns>
        private CellFormats CreateCellFormats()
        {
            var cellFormats = new CellFormats { Count = 3 };
            var alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left };

            cellFormats.Append(StyleTemplates.GetCellFormat(0, 0, false, (Alignment)alignment.Clone()));
            cellFormats.Append(StyleTemplates.GetCellFormat(14, 0, true, (Alignment)alignment.Clone()));
            cellFormats.Append(StyleTemplates.GetCellFormat(0, 1, true, (Alignment)alignment.Clone()));

            for (var index = 0; index < _table.Columns.Count; index++)
            {
                var tableColumn = _table.Columns[index];
                _styleIndex[tableColumn] = cellFormats.Count;

                var cellFormat = StyleTemplates.GetCellFormat(0, 0);

                if (tableColumn.Format != null)
                {
                    cellFormat.NumberFormatId = NUM_FMT_ID_INDENT + (uint)index;
                    cellFormat.ApplyNumberFormat = true;
                }

                cellFormats.Append(cellFormat);
                cellFormats.Count++;
            }

            return cellFormats;
        }

        /// <summary>
        /// Создание форматов стилей ячеек.
        /// </summary>
        /// <returns>Форматы стиля ячейки.</returns>
        private static CellStyleFormats CreateCellStyleFormats()
        {
            var cellStyleFormats = new CellStyleFormats { Count = 1 };

            cellStyleFormats.Append(StyleTemplates.GetCellFormat(0, 0));

            return cellStyleFormats;
        }

        /// <summary>
        /// Создание стилей ячеек.
        /// </summary>
        /// <returns>Стили ячеек.</returns>
        private static CellStyles CreateCellStyles()
        {
            var cellStyles = new CellStyles { Count = 1 };
            cellStyles.Append(StyleTemplates.GetEmptyCellStyle());

            return cellStyles;
        }

        /// <summary>
        /// Создание дифференциальных форматов.
        /// </summary>
        /// <returns>Дифференциальные форматы.</returns>
        private static DifferentialFormats CreateDifferentialFormats() => new DifferentialFormats() { Count = 0 };

        /// <summary>
        /// Создание заливок.
        /// </summary>
        /// <returns>Заливки.</returns>
        private static Fills CreateFills()
        {
            var fills = new Fills { Count = 1 };
            var patternFill = new PatternFill { PatternType = PatternValues.None };
            var fill = new Fill(patternFill);
            fills.Append(fill);

            return fills;
        }

        /// <summary>
        /// Создание шрифтов.
        /// </summary>
        /// <returns>Шрифты.</returns>
        private static Fonts CreateFonts()
        {
            var fonts = new Fonts { Count = 2, KnownFonts = true };

            // Текст
            fonts.Append(StyleTemplates.GetFont());

            // Заголовки
            fonts.Append(StyleTemplates.GetFont(new Bold()));

            return fonts;
        }

        /// <summary>
        /// Создание форматов нумерации.
        /// </summary>
        /// <returns>Форматы нумерации.</returns>
        private NumberingFormats CreateNumeringFormats()
        {
            var numberingFormats = new NumberingFormats { Count = 0 };

            for (var index = 0; index < _table.Columns.Count; index++)
            {
                var tableColumn = _table.Columns[index];

                if (tableColumn.Format != null)
                {
                    numberingFormats.Count++;
                    var numberingFormat = new NumberingFormat
                    {
                        NumberFormatId = NUM_FMT_ID_INDENT + (uint)index,
                        FormatCode = tableColumn.Format
                    };

                    numberingFormats.Append(numberingFormat);
                }
            }
            return numberingFormats;
        }

        /// <summary>
        /// Создание списка расширений таблицы стилей.
        /// </summary>
        /// <returns>Список расширений таблицы стилей.</returns>
        private static StylesheetExtensionList CreateStylesheetExtensionList() => new StylesheetExtensionList();

        /// <summary>
        /// Создание стилей таблицы.
        /// </summary>
        /// <returns>Стили таблицы.</returns>
        private static TableStyles CreateTableStyles() => new TableStyles()
        {
            Count = 0,
            DefaultTableStyle = "TableStyleMedium2",
            DefaultPivotStyle = "PivotStyleLight16"
        };
    }
}