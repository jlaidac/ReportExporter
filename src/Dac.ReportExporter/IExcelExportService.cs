namespace Dac.ReportExporter
{
    public interface IExcelExportService
    {
        Dictionary<CellFormatType, uint> StyleIndices { get; set; }

        MemoryStream ExportToExcelStream(ExportDataModel exportData, IEnumerable<Styles.IExcelStyleApplier>? styleAppliers = null);
        void AppendRowsToExcelStream(MemoryStream memoryStream, ExportSheetModel sheetModel);
    }
}
