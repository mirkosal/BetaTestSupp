// File: Core/Help.WorkbookWriter.cs
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OxSs = DocumentFormat.OpenXml.Spreadsheet;

namespace BetaTestSupp.Core
{
    public interface IWorkbookWriter
    {
        // Crea un XLSX con un solo foglio e scrive header + rows
        void WriteTable(string path, string sheetName, IList<string> headers, IEnumerable<IList<string>> rows);
    }

    public sealed class HelpWorkbookWriter : IWorkbookWriter
    {
        public void WriteTable(string path, string sheetName, IList<string> headers, IEnumerable<IList<string>> rows)
        {
            using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new OxSs.Workbook();
            var sheets = wbPart.Workbook.AppendChild(new OxSs.Sheets());

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new OxSs.Worksheet(new OxSs.SheetData());
            wsPart.Worksheet.Save();

            var relId = wbPart.GetIdOfPart(wsPart);
            sheets.Append(new OxSs.Sheet { Id = relId, SheetId = 1U, Name = string.IsNullOrWhiteSpace(sheetName) ? "Fase3" : sheetName });

            var sd = wsPart.Worksheet.GetFirstChild<OxSs.SheetData>()!;

            // Header
            sd.AppendChild(MakeRow(headers));

            // Rows
            foreach (var r in rows)
                sd.AppendChild(MakeRow(r));

            wbPart.Workbook.Save();
        }

        private static OxSs.Row MakeRow(IList<string> cells)
        {
            var row = new OxSs.Row();
            foreach (var c in cells)
                row.AppendChild(new OxSs.Cell
                {
                    DataType = OxSs.CellValues.String,
                    CellValue = new OxSs.CellValue(c ?? string.Empty)
                });
            return row;
        }
    }
}
