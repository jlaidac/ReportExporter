using Dac.ReportExporter.Styles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dac.ReportExporter
{
    public class ExcelExportService : IExcelExportService
    {
        private readonly IExcelStyleApplier _defaultExcelStyleApplier;

        public ExcelExportService(IExcelStyleApplier defaultExcelStyleApplier)
        {
            _defaultExcelStyleApplier = defaultExcelStyleApplier;
        }

        /// <summary>
        /// Dictionary to store the style indices for each cell format type.
        /// You can manually assign this dictionary for each CellFormatType and the StyleIndex when assign style by StylesheetXml
        /// </summary>
        public Dictionary<CellFormatType, uint> StyleIndices { get; set; } = new Dictionary<CellFormatType, uint>();

        public MemoryStream ExportToExcelStream(ExportDataModel exportData, IEnumerable<IExcelStyleApplier>? styleAppliers = null)
        {
            MemoryStream memoryStream = new MemoryStream();
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook, true))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                CreateWorkbook(workbookPart, exportData, styleAppliers);
            }

            memoryStream.Position = 0; // Reset the position to the beginning of the stream
            return memoryStream;
        }

        public void AppendRowsToExcelStream(MemoryStream memoryStream, ExportSheetModel sheetModel)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(memoryStream, true))
            {
                WorkbookPart? workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    throw new InvalidOperationException("The workbook part is missing from the document.");
                }

                if (workbookPart.Workbook.Sheets == null)
                {
                    throw new InvalidOperationException("The sheets are missing from the workbook.");
                }

                Sheet? sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetModel.SheetName);
                if (sheet == null || sheet.Id == null)
                {
                    throw new InvalidOperationException($"The sheet with name '{sheetModel.SheetName}' does not exist in the workbook.");
                }
                WorksheetPart replacementPart = workbookPart.AddNewPart<WorksheetPart>();

                WorksheetPart existingWorksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

                // Append the new rows to the sheet using OpenXmlWriter
                AppendRows(existingWorksheetPart, replacementPart, sheetModel);

                string replacementPartId = workbookPart.GetIdOfPart(replacementPart);
                sheet.Id.Value = replacementPartId;

                workbookPart.DeletePart(existingWorksheetPart);

                workbookPart.Workbook.Save();
            }

            memoryStream.Position = 0; // Reset the position to the beginning of the stream
        }

        private void CreateWorkbook(WorkbookPart workbookPart, ExportDataModel exportData, IEnumerable<IExcelStyleApplier>? styleAppliers)
        {
            workbookPart.Workbook = new Workbook();

            var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet.Save();

            _defaultExcelStyleApplier.ApplyStyle(workbookStylesPart);

            // Apply cell styles to the workbook part and store style indices
            ApplyCustomizedStyles(workbookStylesPart, styleAppliers);

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
            uint sheetId = 1;
            foreach (var sheetModel in exportData.Sheets)
            {
                sheetModel.SheetId = sheetId++;
                CreateSheet(workbookPart, sheets, sheetModel);
            }

            workbookPart.Workbook.Save();
        }

        private Dictionary<CellFormatType, uint> ApplyCustomizedStyles(WorkbookStylesPart workbookStylesPart, IEnumerable<IExcelStyleApplier>? styleAppliers)
        {
            if (styleAppliers is not null)
            {
                foreach (var styleApplier in styleAppliers)
                {
                    StyleIndices[styleApplier.CellFormatType] = styleApplier.ApplyStyle(workbookStylesPart);
                }
            }

            return StyleIndices;
        }

        private void CreateSheet(WorkbookPart workbookPart, Sheets sheets, ExportSheetModel sheetModel)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();

            // Create custom widths for columns
            if (!string.IsNullOrEmpty(sheetModel.ColumnsXml))
            {
                Columns columns = new Columns();
                columns.InnerXml = sheetModel.ColumnsXml;
                worksheetPart.Worksheet.Append(columns);
                worksheetPart.Worksheet.Save();
            }

            Sheet sheetElement = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetModel.SheetId,
                Name = sheetModel.SheetName
            };
            sheets.Append(sheetElement);

            AddSheetData(worksheetPart, sheetModel);
        }

        private void AddSheetData(WorksheetPart worksheetPart, ExportSheetModel sheetModel)
        {
            using (OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart))
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                var headerRow = new Row();
                AppendHeaders(headerRow, sheetModel);
                writer.WriteElement(headerRow);

                WriteNewRows(writer, sheetModel);

                writer.WriteEndElement(); // End SheetData
                writer.WriteEndElement(); // End Worksheet
            }
        }

        private void AppendRows(WorksheetPart worksheetPart, WorksheetPart replacementPart, ExportSheetModel sheetModel)
        {
            using (OpenXmlWriter writer = OpenXmlWriter.Create(replacementPart))
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                using (OpenXmlReader reader = OpenXmlReader.Create(worksheetPart))
                {
                    WriteExistingRows(reader, writer, sheetModel);
                }

                WriteNewRows(writer, sheetModel);

                writer.WriteEndElement(); // End SheetData
                writer.WriteEndElement(); // End Worksheet
            }
        }

        private void WriteExistingRows(OpenXmlReader reader, OpenXmlWriter writer, ExportSheetModel sheetModel)
        {
            bool headerRowFound = false;
            while (reader.Read())
            {
                if (reader.ElementType == typeof(Row))
                {
                    Row currentRow = (Row)reader.LoadCurrentElement()!;
                    if (headerRowFound)
                    {
                        writer.WriteElement(currentRow);
                    }
                    else
                    {
                        headerRowFound = true;
                        AppendHeaders(currentRow, sheetModel);
                        writer.WriteElement(currentRow);
                    }
                }
            }
        }

        private void AppendHeaders(Row currentRow, ExportSheetModel sheetModel)
        {
            int existingHeaderCount = currentRow.Elements<Cell>().Count();
            int newHeaderCount = sheetModel.Headers.Count;
            if (existingHeaderCount < newHeaderCount)
            {
                for (int i = existingHeaderCount; i < newHeaderCount; i++)
                {
                    Cell newHeaderCell = CreateStyledCell(sheetModel.Headers[i], sheetModel.HeaderFormat);
                    currentRow.AppendChild(newHeaderCell);
                }
            }
        }

        private void WriteNewRows(OpenXmlWriter writer, ExportSheetModel sheetModel)
        {
            foreach (var rowModel in sheetModel.Rows)
            {
                writer.WriteStartElement(new Row());
                foreach (var cellModel in rowModel.Cells)
                {
                    Cell cell = CreateStyledCell(cellModel.Value, cellModel.CellFormat);
                    writer.WriteElement(cell);
                }
                writer.WriteEndElement(); // End Row
            }
        }

        private Cell CreateStyledCell(string value, CellFormatType? formatType)
        {
            Cell cell = new Cell() { CellValue = new CellValue(value), DataType = CellValues.String };
            if (formatType.HasValue && StyleIndices.ContainsKey(formatType.Value))
            {
                cell.StyleIndex = StyleIndices[formatType.Value];
            }
            return cell;
        }
    }
}