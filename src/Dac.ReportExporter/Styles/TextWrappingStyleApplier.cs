using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dac.ReportExporter.Styles
{
    public class TextWrappingStyleApplier : IExcelStyleApplier
    {
        public CellFormatType CellFormatType => CellFormatType.WrapText;

        public uint ApplyStyle(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = workbookStylesPart.Stylesheet;

            CellFormat cellFormat = new CellFormat()
            {
                ApplyAlignment = true,
                Alignment = new Alignment() { WrapText = true }
            };

            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats();
            }

            stylesheet.CellFormats!.Append(cellFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;

            stylesheet.Save();

            return (uint)stylesheet.CellFormats.ChildElements.Count - 1;
        }
    }
}