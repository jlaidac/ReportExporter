using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dac.ReportExporter.Styles
{
    public class DefaultStyleApplier : IExcelStyleApplier
    {
        public CellFormatType CellFormatType => CellFormatType.Undefined;

        public uint ApplyStyle(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = workbookStylesPart.Stylesheet;

            // Create "fonts" node.
            var fonts = new Fonts();
            fonts.Append(new Font()
            {
                FontName = new FontName() { Val = "Calibri" },
                FontSize = new FontSize() { Val = 11 },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
            });
            fonts.Count = (uint)fonts.ChildElements.Count;
            stylesheet.Append(fonts);

            // Create "fills" node.
            var fills = new Fills();
            fills.Append(new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            });
            fills.Count = (uint)fills.ChildElements.Count;
            stylesheet.Append(fills);

            // Create "borders" node.
            var borders = new Borders();
            borders.Append(new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = (uint)borders.ChildElements.Count;
            stylesheet.Append(borders);

            // Create "cellStyleXfs" node.
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.Append(new CellFormat()
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });
            cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;
            stylesheet.Append(cellStyleFormats);

            // Create "cellXfs" node.
            // A default style that works for everything but DateTime
            var cellFormats = new CellFormats();
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            });
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;
            stylesheet.Append(cellFormats);

            stylesheet.Save();
            return 0;
        }
    }
}