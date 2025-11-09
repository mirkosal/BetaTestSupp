// File: Core/Help.ListSource.cs
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using OxSs = DocumentFormat.OpenXml.Spreadsheet;

namespace BetaTestSupp.Core
{
    public sealed class HelpListSourceReader : IListSourceReader
    {
        public List<string> GetSheets(string path)
        {
            var ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            if (ext != ".xlsx") return new();
            using var doc = SpreadsheetDocument.Open(path, false);
            return doc.WorkbookPart!.Workbook.Sheets!.Elements<OxSs.Sheet>()
                     .Select(s => s.Name?.Value ?? "").Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        }

        public List<string> GetHeaders(string path, string? sheetName)
        {
            var ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            if (ext == ".xlsx") return GetHeadersXlsx(path, sheetName);
            if (ext is ".csv" or ".txt") return GetHeadersDelimited(path);
            return new(); // non tabellare
        }

        public List<string> ReadValues(string path, string? sheetName, int selectedHeaderIndex0)
        {
            var ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            if (ext == ".xlsx") return ReadValuesXlsx(path, sheetName, selectedHeaderIndex0);
            if (ext is ".csv" or ".txt") return ReadValuesDelimited(path, selectedHeaderIndex0);
            // fallback: riga = valore
            return File.ReadAllLines(path, new UTF8Encoding(false)).Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        }

        // ---- XLSX helpers ----
        private static List<string> GetHeadersXlsx(string path, string? sheetName)
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var wb = doc.WorkbookPart!.Workbook;
            var sheet = wb.Sheets!.Elements<OxSs.Sheet>()
                         .FirstOrDefault(s => string.Equals(s.Name?.Value ?? "", sheetName ?? "", StringComparison.OrdinalIgnoreCase));
            if (sheet == null) return new();

            var wsp = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id!);
            var sd = wsp.Worksheet.GetFirstChild<OxSs.SheetData>();
            if (sd == null) return new();

            var row1 = sd.Elements<OxSs.Row>().FirstOrDefault(r => r.RowIndex?.Value == 1U);
            if (row1 == null) return new();

            string GetStr(OxSs.Cell c)
            {
                if (c.DataType?.Value == OxSs.CellValues.SharedString)
                {
                    var sst = doc.WorkbookPart.SharedStringTablePart?.SharedStringTable;
                    if (sst == null) return "";
                    if (int.TryParse(c.CellValue?.InnerText, out int idx) && idx >= 0 && idx < sst.Count())
                        return sst.ElementAt(idx).InnerText ?? "";
                    return "";
                }
                return c.CellValue?.InnerText ?? "";
            }

            static string ExcelColName(int index) { int n = index; string s = ""; while (n > 0) { n--; s = (char)('A' + (n % 26)) + s; n /= 26; } return s; }
            static int ColIndex(string a1) { int i = 0; foreach (var ch in a1) { if (char.IsLetter(ch)) i = i * 26 + (char.ToUpperInvariant(ch) - 'A' + 1); else break; } return i; }

            var list = row1.Elements<OxSs.Cell>()
                           .Select(c => (ColIndex(c.CellReference!.Value), GetStr(c)))
                           .OrderBy(t => t.Item1)
                           .Select((t, i) => $"{ExcelColName(t.Item1)}: {(string.IsNullOrWhiteSpace(t.Item2) ? "(senza nome)" : t.Item2.Trim())}")
                           .ToList();
            return list;
        }

        private static List<string> ReadValuesXlsx(string path, string? sheetName, int headerIdx0)
        {
            var res = new List<string>();
            using var doc = SpreadsheetDocument.Open(path, false);
            var wb = doc.WorkbookPart!.Workbook;
            var sheet = wb.Sheets!.Elements<OxSs.Sheet>()
                         .FirstOrDefault(s => string.Equals(s.Name?.Value ?? "", sheetName ?? "", StringComparison.OrdinalIgnoreCase));
            if (sheet == null) return res;

            var wsp = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id!);
            var sd = wsp.Worksheet.GetFirstChild<OxSs.SheetData>();
            if (sd == null) return res;

            int headerCol = headerIdx0 + 1;
            static string A1(uint row, int col)
            {
                string name = ""; int n = col; while (n > 0) { n--; name = (char)('A' + (n % 26)) + name; n /= 26; }
                return name + row;
            }
            string GetStr(OxSs.Cell c)
            {
                if (c.DataType?.Value == OxSs.CellValues.SharedString)
                {
                    var sst = doc.WorkbookPart.SharedStringTablePart?.SharedStringTable;
                    if (sst == null) return "";
                    if (int.TryParse(c.CellValue?.InnerText, out int idx) && idx >= 0 && idx < sst.Count())
                        return sst.ElementAt(idx).InnerText ?? "";
                    return "";
                }
                return c.CellValue?.InnerText ?? "";
            }

            foreach (var row in sd.Elements<OxSs.Row>())
            {
                if (row.RowIndex!.Value < 2U) continue; // salta header
                var cellRef = A1(row.RowIndex!.Value, headerCol);
                var cell = row.Elements<OxSs.Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value ?? "", cellRef, StringComparison.OrdinalIgnoreCase));
                if (cell == null) continue;
                var v = GetStr(cell).Trim();
                if (!string.IsNullOrWhiteSpace(v)) res.Add(v);
            }
            return res;
        }

        // ---- CSV/TXT helpers ----
        private static List<string> GetHeadersDelimited(string path)
        {
            var lines = File.ReadAllLines(path, new UTF8Encoding(false));
            if (lines.Length == 0) return new();
            char sep = GuessSep(lines[0]);
            var headers = lines[0].Split(sep).Select((h, i) => $"{ExcelColName(i + 1)}: {(string.IsNullOrWhiteSpace(h) ? "(senza nome)" : h.Trim())}").ToList();
            return headers;

            static char GuessSep(string line) => new[] { ';', ',', '\t', '|' }.OrderByDescending(c => line.Count(ch => ch == c)).First();
            static string ExcelColName(int index) { int n = index; string s = ""; while (n > 0) { n--; s = (char)('A' + (n % 26)) + s; n /= 26; } return s; }
        }

        private static List<string> ReadValuesDelimited(string path, int headerIdx0)
        {
            var res = new List<string>();
            var lines = File.ReadAllLines(path, new UTF8Encoding(false));
            if (lines.Length <= 1) return res;
            char sep = new[] { ';', ',', '\t', '|' }.OrderByDescending(c => lines[0].Count(ch => ch == c)).First();

            for (int i = 1; i < lines.Length; i++)
            {
                var parts = lines[i].Split(sep);
                if (headerIdx0 < parts.Length)
                {
                    var v = (parts[headerIdx0] ?? "").Trim();
                    if (!string.IsNullOrWhiteSpace(v)) res.Add(v);
                }
            }
            return res;
        }
    }
}
