using DocumentFormat.OpenXml.Packaging;

namespace Dac.ReportExporter.Styles
{
    public interface IExcelStyleApplier
    {
        CellFormatType CellFormatType { get; }
        uint ApplyStyle(WorkbookStylesPart workbookStylesPart);
    }
}