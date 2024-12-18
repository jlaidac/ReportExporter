namespace Dac.ReportExporter
{
    public class ExportDataModel
    {
        public string? FileName { get; set; }
        public List<ExportSheetModel> Sheets { get; set; } = new();
    }

    public class ExportSheetModel
    {
        public uint SheetId { get; set; } // 1-based
        public string SheetName { get; set; } = string.Empty;
        public CellFormatType? HeaderFormat { get; set; }
        public string? ColumnsXml { get; set; }

        public List<string> Headers { get; set; } = new List<string>();
        public List<RowModel> Rows { get; set; } = new List<RowModel>();
    }

    public class RowModel
    {
        public List<CellModel> Cells { get; set; } = new();
    }

    public struct CellModel
    {
        public string Value { get; set; }
        public CellFormatType? CellFormat { get; set; }
    }

    public enum CellFormatType
    {
        Undefined = 0,
        WrapText,
        BoldText
    }
}
