using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

// OpenXML
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public partial class CsvToXlsxForm : BaseMenuForm
{
    public CsvToXlsxForm()
    {
        InitializeComponent();

        // Se hai due TextBox per separatore e escape, limita a un char
        txtSep.TextChanged += EnforceOneChar;
        txtEsc.TextChanged += EnforceOneChar;

        btnBrowse.Click += btnBrowse_Click;
        btnConvert.Click += btnConvert_Click;
    }

    private void btnBrowse_Click(object? sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog
        {
            Title = "Seleziona un file CSV",
            Filter = "CSV (*.csv)|*.csv|Tutti i file (*.*)|*.*",
            Multiselect = false,
            CheckFileExists = true
        };
        if (ofd.ShowDialog(this) == DialogResult.OK)
            txtCsvPath.Text = ofd.FileName;
    }

    private void EnforceOneChar(object? sender, EventArgs e)
    {
        if (sender is TextBox t && t.Text.Length > 1)
            t.Text = t.Text.Substring(0, 1);
    }

    private void btnConvert_Click(object? sender, EventArgs e)
    {
        var csvPath = (txtCsvPath.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(csvPath) || !File.Exists(csvPath))
        {
            MessageBox.Show(this, "Seleziona un file CSV valido.", "Attenzione",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        char sep = (txtSep.Text?.Length == 1) ? txtSep.Text![0] : ';';
        char esc = (txtEsc.Text?.Length == 1) ? txtEsc.Text![0] : '"';
        bool hasHeader = chkHeader.Checked;
        bool detectAllDates = chkDetectAllDates.Checked;

        using var sfd = new SaveFileDialog
        {
            Title = "Salva come XLSX",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = Path.GetFileNameWithoutExtension(csvPath) + ".xlsx"
        };
        if (sfd.ShowDialog(this) != DialogResult.OK) return;

        try
        {
            var rows = ParseCsv(csvPath, sep, esc);
            if (rows.Count == 0)
            {
                MessageBox.Show(this, "Il CSV è vuoto.", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var doc = SpreadsheetDocument.Create(sfd.FileName, SpreadsheetDocumentType.Workbook))
            {
                var wbp = doc.AddWorkbookPart();
                wbp.Workbook = new Workbook();
                var sheets = wbp.Workbook.AppendChild(new Sheets());

                var wsp = wbp.AddNewPart<WorksheetPart>();
                wsp.Worksheet = new Worksheet(new SheetData());
                wsp.Worksheet.Save();

                var relId = wbp.GetIdOfPart(wsp);
                uint sid = 1;
                sheets.Append(new Sheet() { Id = relId, SheetId = sid, Name = "Sheet1" });

                var sd = wsp.Worksheet.GetFirstChild<SheetData>()!;
                int rindex = 1;

                foreach (var row in rows)
                {
                    var r = new Row() { RowIndex = (uint)rindex };
                    sd.Append(r);

                    for (int c = 1; c <= row.Count; c++)
                    {
                        string raw = row[c - 1] ?? "";
                        bool isHeaderRow = hasHeader && rindex == 1;

                        string text = raw;
                        if (!isHeaderRow && detectAllDates && TryParseAnyDate(raw, out DateTime dt))
                            text = dt.ToString("yyyy-MM-dd");

                        var cell = new Cell()
                        {
                            CellReference = RefA1(rindex, c),
                            DataType = CellValues.String,
                            CellValue = new CellValue(text)
                        };
                        r.Append(cell);
                    }
                    rindex++;
                }

                wsp.Worksheet.Save();
                wbp.Workbook.Save();
            }

            MessageBox.Show(this, "Conversione completata.", "OK",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Errore durante la conversione:\n" + ex.Message, "Errore",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static List<List<string>> ParseCsv(string path, char sep, char esc)
    {
        var lines = new List<List<string>>();
        using var sr = new StreamReader(path);

        string? line;
        while ((line = sr.ReadLine()) != null)
        {
            var row = new List<string>();
            var cur = new System.Text.StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];

                if (inQuotes)
                {
                    if (ch == esc)
                    {
                        bool doubled = (i + 1 < line.Length && line[i + 1] == esc);
                        if (doubled)
                        {
                            cur.Append(esc);
                            i++;
                            continue;
                        }
                        inQuotes = false;
                        continue;
                    }
                    if (ch == '\\' && i + 1 < line.Length && line[i + 1] == esc)
                    {
                        cur.Append(esc);
                        i++;
                        continue;
                    }
                    cur.Append(ch);
                }
                else
                {
                    if (ch == esc)
                    {
                        inQuotes = true;
                    }
                    else if (ch == sep)
                    {
                        row.Add(cur.ToString());
                        cur.Clear();
                    }
                    else
                    {
                        cur.Append(ch);
                    }
                }
            }
            row.Add(cur.ToString());
            lines.Add(row);
        }

        return lines;
    }

    private static bool TryParseAnyDate(string text, out DateTime dt)
    {
        text = (text ?? "").Trim();
        if (text.Length == 0) { dt = default; return false; }

        string[] fmts = new[]
        {
            "yyyy-MM-dd","yyyy/MM/dd","dd/MM/yyyy","dd-MM-yyyy","MM/dd/yyyy","MM-dd-yyyy",
            "yyyyMMdd","ddMMyyyy","yyyy.MM.dd","dd.MM.yyyy",
            "yyyy-MM-ddTHH:mm:ss","yyyy-MM-dd HH:mm:ss","dd/MM/yyyy HH:mm:ss"
        };

        if (DateTime.TryParseExact(text, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return true;
        if (DateTime.TryParseExact(text, fmts, new CultureInfo("it-IT"), DateTimeStyles.None, out dt)) return true;

        if (DateTime.TryParse(text, new CultureInfo("it-IT"), DateTimeStyles.AssumeLocal, out dt)) return true;
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dt)) return true;

        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double oa))
        {
            try { dt = DateTime.FromOADate(oa); return true; } catch { /* ignore */ }
        }

        dt = default;
        return false;
    }

    private static string RefA1(int row, int col)
    {
        string colName = "";
        int n = col;
        while (n > 0) { n--; colName = (char)('A' + (n % 26)) + colName; n /= 26; }
        return $"{colName}{row}";
    }
}
