using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

// OpenXML
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// Evita ambiguità con Spreadsheet.Color
using WinColor = System.Drawing.Color;

public partial class CourseCompareForm : BaseMenuForm
{
    public CourseCompareForm()
    {
        InitializeComponent();

        // Wire-up
        btnBrowseA.Click += (_, __) => BrowseFile(txtFileA);
        btnBrowseB.Click += (_, __) => BrowseFile(txtFileB);
        btnLoadHeaders.Click += (_, __) => LoadHeaders();
        btnRun.Click += (_, __) => RunComparison();
    }

    // ======== UI helpers ========
    private void BrowseFile(TextBox target)
    {
        using var ofd = new OpenFileDialog
        {
            Title = "Seleziona un file Excel (.xlsx)",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            CheckFileExists = true
        };
        if (ofd.ShowDialog(this) == DialogResult.OK)
        {
            target.Text = ofd.FileName;
            lblStatus.Text = $"Selezionato: {ofd.FileName}";
        }
    }

    private class ColItem
    {
        public string Text { get; set; } = "";
        public int Index { get; set; }  // 1-based
        public override string ToString() => Text;
    }

    private void LoadHeaders()
    {
        try
        {
            lstACol1.Items.Clear(); lstACol2.Items.Clear();
            lstBCol1.Items.Clear(); lstBCol2.Items.Clear();

            var pathA = (txtFileA.Text ?? "").Trim();
            var pathB = (txtFileB.Text ?? "").Trim();
            if (!File.Exists(pathA) && !File.Exists(pathB))
            {
                MessageBox.Show(this, "Seleziona almeno un file (meglio entrambi) prima di caricare le intestazioni.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (File.Exists(pathA))
            {
                var headersA = ReadHeaders(pathA);
                for (int c = 1; c <= headersA.Count; c++)
                {
                    string letter = ExcelColName(c);
                    string h = string.IsNullOrWhiteSpace(headersA[c - 1]) ? "(senza nome)" : headersA[c - 1];
                    var item = new ColItem { Text = $"{letter}: {h}", Index = c };
                    lstACol1.Items.Add(item);
                    lstACol2.Items.Add(new ColItem { Text = item.Text, Index = c });
                }
                if (lstACol1.Items.Count > 0) lstACol1.SelectedIndex = 0;
                if (lstACol2.Items.Count > 1) lstACol2.SelectedIndex = 1;
            }

            if (File.Exists(pathB))
            {
                var headersB = ReadHeaders(pathB);
                for (int c = 1; c <= headersB.Count; c++)
                {
                    string letter = ExcelColName(c);
                    string h = string.IsNullOrWhiteSpace(headersB[c - 1]) ? "(senza nome)" : headersB[c - 1];
                    var item = new ColItem { Text = $"{letter}: {h}", Index = c };
                    lstBCol1.Items.Add(item);
                    lstBCol2.Items.Add(new ColItem { Text = item.Text, Index = c });
                }
                if (lstBCol1.Items.Count > 0) lstBCol1.SelectedIndex = 0;
                if (lstBCol2.Items.Count > 1) lstBCol2.SelectedIndex = 1;
            }

            lblStatus.Text = "Intestazioni caricate.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Errore durante la lettura intestazioni:\n" + ex.Message, "Errore",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ======== Confronto e salvataggio ========
    private void RunComparison()
    {
        try
        {
            var pathA = (txtFileA.Text ?? "").Trim();
            var pathB = (txtFileB.Text ?? "").Trim();

            if (!File.Exists(pathA) || !File.Exists(pathB))
            {
                MessageBox.Show(this, "Seleziona sia il File A sia il File B.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (lstACol1.SelectedItem is not ColItem a1 || lstACol2.SelectedItem is not ColItem a2 ||
                lstBCol1.SelectedItem is not ColItem b1 || lstBCol2.SelectedItem is not ColItem b2)
            {
                MessageBox.Show(this, "Seleziona due colonne per il File A e due colonne per il File B.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (a1.Index == a2.Index || b1.Index == b2.Index)
            {
                MessageBox.Show(this, "Le due colonne scelte per ciascun file devono essere diverse.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool outputFromA = rdoOutFromA.Checked; // se false -> da B

            using var sfd = new SaveFileDialog
            {
                Title = "Scegli dove salvare l'output",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = outputFromA ? "Differenze_da_A.xlsx" : "Differenze_da_B.xlsx"
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            // Leggi i dati completi dei due file (prima sheet)
            var (headersA, rowsA) = ReadSheetAsMatrix(pathA);
            var (headersB, rowsB) = ReadSheetAsMatrix(pathB);

            // Costruisci set chiavi composte
            var keysA = BuildKeySet(rowsA, a1.Index, a2.Index);
            var keysB = BuildKeySet(rowsB, b1.Index, b2.Index);

            // Scegli sorgente di output
            var sourceHeaders = outputFromA ? headersA : headersB;
            var sourceRows = outputFromA ? rowsA : rowsB;
            var otherKeys = outputFromA ? keysB : keysA;

            // Filtra righe la cui chiave composta NON esiste nell'altro set
            var toWrite = new List<string[]>();
            foreach (var r in sourceRows)
            {
                string k = MakeKey(
                    SafeGet(r, outputFromA ? a1.Index : b1.Index),
                    SafeGet(r, outputFromA ? a2.Index : b2.Index));
                if (!otherKeys.Contains(k))
                    toWrite.Add(r);
            }

            // Scrivi un unico foglio "Output" con header del file scelto e tutte le righe
            WriteXlsxSingleSheet(sfd.FileName, "Output", sourceHeaders, toWrite);

            lblStatus.Text = $"Fatto. Salvato: {sfd.FileName}";
            MessageBox.Show(this, $"Operazione completata.\nRighe scritte: {toWrite.Count}", "OK",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Errore durante il confronto/salvataggio:\n" + ex.Message, "Errore",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ======== Lettura con OpenXML (primo foglio) ========
    private static (List<string> headers, List<string[]> rows) ReadSheetAsMatrix(string path)
    {
        var headers = ReadHeaders(path);
        var rows = new List<string[]>();

        using var doc = SpreadsheetDocument.Open(path, false);
        var wbp = doc.WorkbookPart!;
        var sheet = wbp.Workbook.Sheets!.Elements<Sheet>().First();
        var wsp = (WorksheetPart)wbp.GetPartById(sheet.Id!);
        var sd = wsp.Worksheet.GetFirstChild<SheetData>() ?? new SheetData();

        int maxCols = headers.Count;
        foreach (var row in sd.Elements<Row>())
        {
            if (row.RowIndex == null) continue;
            int rIndex = (int)row.RowIndex.Value;
            if (rIndex == 1) continue; // salta header

            var arr = new string[maxCols];
            foreach (var cell in row.Elements<Cell>())
            {
                int c = ColumnIndex(cell.CellReference!);
                if (c >= 1 && c <= maxCols)
                    arr[c - 1] = GetCellText(cell, wbp);
            }
            rows.Add(arr);
        }

        return (headers, rows);
    }

    private static List<string> ReadHeaders(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wbp = doc.WorkbookPart!;
        var sheet = wbp.Workbook.Sheets!.Elements<Sheet>().First();
        var wsp = (WorksheetPart)wbp.GetPartById(sheet.Id!);
        var sd = wsp.Worksheet.GetFirstChild<SheetData>() ?? new SheetData();

        var headers = new List<string>();
        var headerRow = sd.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == 1);
        if (headerRow != null)
        {
            int maxCol = 0;
            foreach (var c in headerRow.Elements<Cell>())
                maxCol = Math.Max(maxCol, ColumnIndex(c.CellReference!));

            if (maxCol == 0) maxCol = headerRow.Elements<Cell>().Count();

            for (int col = 1; col <= maxCol; col++)
            {
                var a1 = RefA1(1, col);
                var cell = headerRow.Elements<Cell>().FirstOrDefault(cc => cc.CellReference == a1);
                string text = cell != null ? GetCellText(cell, wbp) : "";
                headers.Add(text ?? "");
            }
        }
        return headers;
    }

    private static string GetCellText(Cell cell, WorkbookPart wbp)
    {
        if (cell == null) return "";
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var sst = wbp.SharedStringTablePart?.SharedStringTable;
            if (sst == null) return "";
            if (!int.TryParse(cell.CellValue?.InnerText ?? "0", out int idx)) return "";
            return sst.ElementAt(idx).InnerText ?? "";
        }
        return cell.CellValue?.InnerText ?? "";
    }

    private static HashSet<string> BuildKeySet(List<string[]> rows, int col1Index1Based, int col2Index1Based)
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        foreach (var r in rows)
        {
            string k = MakeKey(SafeGet(r, col1Index1Based), SafeGet(r, col2Index1Based));
            set.Add(k);
        }
        return set;
    }

    private static string MakeKey(string v1, string v2)
        => (v1 ?? "") + "\u001F" + (v2 ?? ""); // separatore non stampabile per evitare collisioni

    private static string SafeGet(string[] row, int index1Based)
    {
        int idx = Math.Max(1, index1Based) - 1;
        if (idx < 0 || idx >= row.Length) return "";
        return row[idx] ?? "";
    }

    // ======== Scrittura XLSX (un unico foglio) ========
    private static void WriteXlsxSingleSheet(string path, string sheetName, List<string> headers, List<string[]> rows)
    {
        if (File.Exists(path)) File.Delete(path);

        using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            var wbp = doc.AddWorkbookPart();
            wbp.Workbook = new Workbook();
            var sheets = wbp.Workbook.AppendChild(new Sheets());

            var wsp = wbp.AddNewPart<WorksheetPart>();
            wsp.Worksheet = new Worksheet(new SheetData());
            wsp.Worksheet.Save();

            var relId = wbp.GetIdOfPart(wsp);
            uint sid = 1;
            sheets.Append(new Sheet() { Id = relId, SheetId = sid, Name = sheetName });

            var sd = wsp.Worksheet.GetFirstChild<SheetData>()!;

            // Header
            var r1 = new Row() { RowIndex = 1 };
            for (int c = 1; c <= headers.Count; c++)
            {
                var cell = new Cell
                {
                    CellReference = RefA1(1, c),
                    DataType = CellValues.String,
                    CellValue = new CellValue(headers[c - 1] ?? "")
                };
                r1.Append(cell);
            }
            sd.Append(r1);

            // Righe
            int rIdx = 2;
            foreach (var arr in rows)
            {
                var rr = new Row() { RowIndex = (uint)rIdx };
                for (int c = 1; c <= headers.Count; c++)
                {
                    string v = (c - 1 < arr.Length) ? (arr[c - 1] ?? "") : "";
                    var cell = new Cell
                    {
                        CellReference = RefA1(rIdx, c),
                        DataType = CellValues.String,
                        CellValue = new CellValue(v)
                    };
                    rr.Append(cell);
                }
                sd.Append(rr);
                rIdx++;
            }

            wsp.Worksheet.Save();
            wbp.Workbook.Save();
        }
    }

    // ======== A1 helpers ========
    private static string ExcelColName(int index)
    {
        int n = index;
        string s = "";
        while (n > 0) { n--; s = (char)('A' + (n % 26)) + s; n /= 26; }
        return s;
    }

    private static int ColumnIndex(string cellRef)
    {
        int i = 0;
        foreach (char ch in cellRef)
        {
            if (char.IsLetter(ch)) i = i * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
            else break;
        }
        return i;
    }

    private static string RefA1(int row, int col) => ExcelColName(col) + row.ToString();
}
