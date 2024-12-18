using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dac.ReportExporter.Styles
{
    public class BoldTextStyleApplier : IExcelStyleApplier
    {
        public CellFormatType CellFormatType => CellFormatType.BoldText;

        public uint ApplyStyle(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = workbookStylesPart.Stylesheet;

            Font boldFont = new Font()
            {
                Bold = new Bold()
            };

            if (stylesheet.Fonts == null)
            {
                stylesheet.Fonts = new Fonts();
            }

            stylesheet.Fonts.Append(boldFont);
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.ChildElements.Count;

            CellFormat cellFormat = new CellFormat()
            {
                FontId = stylesheet.Fonts.Count - 1,
                ApplyFont = true
            };

            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats();
            }

            stylesheet.CellFormats.Append(cellFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;

            stylesheet.Save();

            return (uint)stylesheet.CellFormats.ChildElements.Count - 1;
        }
    }
}