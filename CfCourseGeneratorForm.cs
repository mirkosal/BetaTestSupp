// File: CfCourseGeneratorForm.cs
using BetaTestSupp.Core;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OxSs = DocumentFormat.OpenXml.Spreadsheet;

public partial class CfCourseGeneratorForm : BaseMenuForm
{
    private readonly HelpContainer _cx;
    private readonly IHelpLogger _log;
    private readonly IFileDialogService _dialogs;
    private readonly ISettingsStore _settings;
    private readonly ITemplateRepository _templates;
    private readonly IListSourceReader _lists;
    private readonly IWorkbookWriter _writer;
    private readonly IRuleRepository _rulesRepo;   // NEW

    private readonly List<ParsedRow> _phase1OkRows = new();

    private sealed class ParsedRow
    {
        public string NomeFileOriginale { get; set; } = "";
        public string Estensione { get; set; } = "";
        public string NomeFileSenzaExt { get; set; } = "";
        public string CodiceFiscale { get; set; } = "";
        public string CodiceCorso { get; set; } = "";
        public string DataCompletamentoRaw { get; set; } = "";
        public string DataCompletamento { get; set; } = "";
        public string? ParseError { get; set; }
    }

    public CfCourseGeneratorForm()
    {
        InitializeComponent();

        _cx = new HelpContainer()
            .Register<IHelpLogger>(new UiListBoxLogger(lstLog))
            .Register<IFileDialogService>(new WinFileDialogService(this))
            .Register<ISettingsStore>(new HelpFileSettingsStore("BetaTestSupp", "CfCourseGenerator.settings"))
            .Register<ITemplateRepository>(new HelpTemplateRepository("BetaTestSupp", "CfCourseGenerator.templates"))
            .Register<IListSourceReader>(new HelpListSourceReader())
            .Register<IWorkbookWriter>(new HelpWorkbookWriter())
            .Register<IRuleRepository>(new HelpRuleRepository("BetaTestSupp", "CfCourseGenerator.rules.json")); // NEW

        _log = _cx.Resolve<IHelpLogger>();
        _dialogs = _cx.Resolve<IFileDialogService>();
        _settings = _cx.Resolve<ISettingsStore>();
        _templates = _cx.Resolve<ITemplateRepository>();
        _lists = _cx.Resolve<IListSourceReader>();
        _writer = _cx.Resolve<IWorkbookWriter>();
        _rulesRepo = _cx.Resolve<IRuleRepository>();

        // === Setup pre-esistente (Fase 1/2/3) ===
        txtCourseMapPath.Text = _settings.Get("Phase2CourseMapPath") ?? "";
        txtPersonMapPath.Text = _settings.Get("Phase2PersonMapPath") ?? "";
        ReloadSavedTemplatesList();
        txtExcel.TextChanged += (_, __) =>
        {
            TryPopulateSheetAndHeaders();
            UpdateDefaultPhase3Out();
            UpdateDefaultPhase4Out();
        };

        // === Fase 4 ===
        btnBrowsePhase4Out.Click += (_, __) => BrowsePhase4Out();
        btnOpenPhase4Folder.Click += (_, __) => OpenPhase4Folder();
        btnReloadFase3Headers.Click += (_, __) => ReloadFase3Headers();
        btnAddRule.Click += (_, __) => AddRule();
        btnEditRule.Click += (_, __) => EditSelectedRule();
        btnDeleteRule.Click += (_, __) => DeleteSelectedRule();
        btnApplyPhase4.Click += (_, __) => ApplyPhase4();

        ReloadRulesList();
        UpdateDefaultPhase4Out();
    }

    // ===== Fase 1/2/3: (metodi già consegnati prima — invariati) =====
    // ... (omessi qui per brevità: BrowseExcel, Generate(), Save/Load template, PreviewPhase3, etc.)

    // =========================
    // ======== Fase 4 =========
    // =========================

    private void UpdateDefaultPhase4Out()
    {
        var inputF3 = (txtPhase3Out.Text ?? "").Trim();
        string baseDir = !string.IsNullOrWhiteSpace(inputF3) && File.Exists(inputF3)
            ? (Path.GetDirectoryName(inputF3) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
            : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        string baseName = !string.IsNullOrWhiteSpace(inputF3) && File.Exists(inputF3)
            ? Path.GetFileNameWithoutExtension(inputF3)
            : "Output_Fase3";

        var suggested = Path.Combine(baseDir, baseName.Replace("_Fase3", "") + "_Fase4_pulito.xlsx");
        if (string.IsNullOrWhiteSpace(txtPhase4Out.Text))
            txtPhase4Out.Text = suggested;
    }

    private void BrowsePhase4Out()
    {
        using var sfd = new SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            Title = "Scegli file di output (Fase 4)",
            FileName = Path.GetFileName(txtPhase4Out.Text)
        };
        if (sfd.ShowDialog(this) == DialogResult.OK)
            txtPhase4Out.Text = sfd.FileName;
    }

    private void OpenPhase4Folder()
    {
        var path = (txtPhase4Out.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(path)) { MessageBox.Show(this, "Seleziona prima un file di output Fase 4."); return; }
        var dir = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir)) { MessageBox.Show(this, "Cartella non trovata."); return; }
        try { Process.Start(new ProcessStartInfo { FileName = dir, UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show(this, $"Impossibile aprire la cartella:\n{ex.Message}"); }
    }

    private void ReloadRulesList()
    {
        chkRules.Items.Clear();
        foreach (var r in _rulesRepo.List())
            chkRules.Items.Add(r, false);
        // visualizza come "Nome regola"
        chkRules.DisplayMember = "Name";
        chkRules.ValueMember = "Id";
    }

    private (List<string> headers, List<List<string>> rows) ReadFase3All()
    {
        var path = (txtPhase3Out.Text ?? "").Trim();
        if (!File.Exists(path))
            throw new FileNotFoundException("File Fase 3 non trovato. Genera prima la Fase 3 o scegli un file valido.", path);

        return ReadSheetAll(path, "Fase3");
    }

    private static (List<string> headers, List<List<string>> rows) ReadSheetAll(string path, string sheetName)
    {
        var headers = new List<string>();
        var rows = new List<List<string>>();

        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart!.Workbook;
        var sheets = wb.Sheets!.Elements<OxSs.Sheet>().ToList();
        var sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value ?? "", sheetName, StringComparison.OrdinalIgnoreCase))
                 ?? sheets.FirstOrDefault();
        if (sheet == null) return (headers, rows);

        var wsp = doc.WorkbookPart!.GetPartById(sheet.Id!) as WorksheetPart;
        var sd = wsp?.Worksheet?.GetFirstChild<OxSs.SheetData>();
        if (sd == null) return (headers, rows);

        string GetStr(OxSs.Cell c)
        {
            try
            {
                if (c.DataType?.Value == OxSs.CellValues.SharedString)
                {
                    var sst = doc.WorkbookPart!.SharedStringTablePart?.SharedStringTable;
                    if (sst == null) return "";
                    if (int.TryParse(c.CellValue?.InnerText, out int idx) && idx >= 0 && idx < sst.Count())
                        return sst.ElementAt(idx).InnerText ?? "";
                    return "";
                }
                return c.CellValue?.InnerText ?? "";
            }
            catch { return ""; }
        }
        static int ColIndex(string a1)
        {
            if (string.IsNullOrEmpty(a1)) return 0;
            int i = 0;
            foreach (var ch in a1)
            {
                if (char.IsLetter(ch)) i = i * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
                else break;
            }
            return i;
        }

        var allRows = sd.Elements<OxSs.Row>().OrderBy(r => r.RowIndex?.Value ?? 0U).ToList();
        if (allRows.Count == 0) return (headers, rows);

        var headerRow = allRows[0];
        var headerCells = headerRow.Elements<OxSs.Cell>().ToList();
        int maxCol = headerCells.Select(c => ColIndex(c.CellReference?.Value ?? "")).DefaultIfEmpty(0).Max();
        if (maxCol == 0) maxCol = headerCells.Count;

        var headerVals = new string[maxCol];
        foreach (var c in headerCells)
        {
            int ci = ColIndex(c.CellReference?.Value ?? "");
            if (ci >= 1 && ci <= maxCol) headerVals[ci - 1] = GetStr(c);
        }
        headers.AddRange(headerVals.Select(s => string.IsNullOrWhiteSpace(s) ? $"Col{Array.IndexOf(headerVals, s) + 1}" : s));

        foreach (var r in allRows.Skip(1))
        {
            var cells = r.Elements<OxSs.Cell>().ToList();
            var vals = new string[maxCol];
            foreach (var c in cells)
            {
                int ci = ColIndex(c.CellReference?.Value ?? "");
                if (ci >= 1 && ci <= maxCol) vals[ci - 1] = GetStr(c);
            }
            rows.Add(vals.ToList());
        }
        return (headers, rows);
    }

    private void ReloadFase3Headers()
    {
        try
        {
            var (headers, _) = ReadFase3All();
            MessageBox.Show(this, "Intestazioni ricaricate:\n- " + string.Join("\n- ", headers), "Fase 3");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Errore nel leggere Fase 3:\n" + ex.Message, "Errore");
        }
    }

    private List<string> GetFase3HeadersOrEmpty()
    {
        try { var (h, _) = ReadFase3All(); return h; }
        catch { return new List<string>(); }
    }

    private void AddRule()
    {
        var hdrs = GetFase3HeadersOrEmpty();
        if (hdrs.Count == 0)
        {
            MessageBox.Show(this, "Genera prima la Fase 3 (o seleziona un file Fase 3) per ottenere le intestazioni.", "Attenzione");
            return;
        }

        using var dlg = new RuleEditorForm(hdrs, null);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _rulesRepo.Save(dlg.Result);
            ReloadRulesList();
        }
    }

    private RuleDef? SelectedRule()
    {
        if (chkRules.SelectedItem is RuleDef r) return r;
        return null;
    }

    private void EditSelectedRule()
    {
        var rule = SelectedRule();
        if (rule == null) { MessageBox.Show(this, "Seleziona una regola.", "Attenzione"); return; }

        var hdrs = GetFase3HeadersOrEmpty();
        using var dlg = new RuleEditorForm(hdrs, rule);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _rulesRepo.Save(dlg.Result);
            ReloadRulesList();
        }
    }

    private void DeleteSelectedRule()
    {
        var rule = SelectedRule();
        if (rule == null) return;
        if (MessageBox.Show(this, $"Eliminare la regola \"{rule.Name}\"?", "Conferma", MessageBoxButtons.YesNo) == DialogResult.Yes)
        {
            _rulesRepo.Delete(rule.Id);
            ReloadRulesList();
        }
    }

    private void ApplyPhase4()
    {
        btnApplyPhase4.Enabled = false;
        pgbPhase4.Visible = true;
        pgbPhase4.Style = ProgressBarStyle.Blocks;

        try
        {
            var selectedRules = chkRules.CheckedItems.Cast<RuleDef>().ToList();
            if (selectedRules.Count == 0)
            {
                MessageBox.Show(this, "Seleziona almeno una regola.", "Attenzione");
                return;
            }

            var (headers, rows) = ReadFase3All();
            if (headers.Count == 0 || rows.Count == 0)
            {
                MessageBox.Show(this, "L'output Fase 3 è vuoto o non leggibile.", "Attenzione");
                return;
            }

            int idxNomeFile = headers.FindIndex(h => string.Equals(h, "NomeFileOriginale", StringComparison.OrdinalIgnoreCase));
            if (idxNomeFile < 0) idxNomeFile = 0; // best effort

            // Prepara set di lookup da Fase 2 (se servono)
            var (peopleSet, peoplePairs) = BuildPhase2PeopleSets();
            var (courseSet, coursePairs) = BuildPhase2CourseSets();

            // Applica regole
            var goodRows = new List<IList<string>>(rows.Count);
            var bad = new List<(string valore, string motivo)>();

            // progress
            pgbPhase4.Maximum = rows.Count; pgbPhase4.Value = 0;

            foreach (var r in rows)
            {
                string? matchRule = MatchesAnyRule(r, headers, selectedRules, peopleSet, peoplePairs, courseSet, coursePairs);
                if (matchRule == null)
                {
                    goodRows.Add(r);
                }
                else
                {
                    string name = (idxNomeFile >= 0 && idxNomeFile < r.Count) ? (r[idxNomeFile] ?? "") : "";
                    bad.Add((name, matchRule));
                }
                if (pgbPhase4.Value < pgbPhase4.Maximum) pgbPhase4.Value++;
                Application.DoEvents();
            }

            // Scrivi Fase4 pulito
            var outPath = (txtPhase4Out.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(outPath)) outPath = SuggestPhase4Name();
            if (!outPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) outPath += ".xlsx";
            _writer.WriteTable(outPath, "Fase4", headers, goodRows); // usa il tuo writer. :contentReference[oaicite:2]{index=2}
            _log.Info($"[Fase4] Generato XLSX pulito: {outPath}");

            // Scrivi CSV esclusi Fase 4
            var exclCsv = Path.Combine(Path.GetDirectoryName(outPath) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "record_esclusi_fase4.csv");
            using (var sw = new StreamWriter(exclCsv, false, new UTF8Encoding(false)))
            {
                sw.WriteLine("ValoreOriginale;Motivo");
                foreach (var e in bad) sw.WriteLine($"{e.valore};{e.motivo}");
            }
            _log.Info($"[Fase4] Esclusi: {bad.Count} → {exclCsv}");

            MessageBox.Show(this, $"Fase 4 completata.\nRighe buone: {goodRows.Count}\nEscluse: {bad.Count}", "OK");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Errore in Fase 4:\n" + ex.Message, "Errore");
        }
        finally
        {
            pgbPhase4.Visible = false;
            btnApplyPhase4.Enabled = true;
        }
    }

    private string SuggestPhase4Name()
    {
        var f3 = (txtPhase3Out.Text ?? "").Trim();
        string dir = File.Exists(f3) ? (Path.GetDirectoryName(f3) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)) : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string baseName = File.Exists(f3) ? Path.GetFileNameWithoutExtension(f3).Replace("_Fase3", "") : "Output";
        return Path.Combine(dir, baseName + "_Fase4_pulito.xlsx");
    }

    private string? MatchesAnyRule(List<string> row, List<string> headers, List<RuleDef> rules,
        HashSet<string> peopleSet, HashSet<(string, string)> peoplePairs,
        HashSet<string> courseSet, HashSet<(string, string)> coursePairs)
    {
        foreach (var rule in rules)
        {
            if (MatchesRule(row, headers, rule, peopleSet, peoplePairs, courseSet, coursePairs))
                return rule.Name; // motivo = nome regola
        }
        return null;
    }

    private bool MatchesRule(List<string> row, List<string> headers, RuleDef rule,
        HashSet<string> peopleSet, HashSet<(string, string)> peoplePairs,
        HashSet<string> courseSet, HashSet<(string, string)> coursePairs)
    {
        string Get(string? fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName)) return "";
            int i = headers.FindIndex(h => string.Equals(h, fieldName, StringComparison.OrdinalIgnoreCase));
            if (i < 0 || i >= row.Count) return "";
            return row[i] ?? "";
        }

        switch (rule.Kind)
        {
            case RuleKind.DateAfterToday:
                {
                    var s = Get(rule.Field1);
                    if (DateTime.TryParse(s, out var dt))
                    {
                        // confronto con oggi (locale)
                        return dt.Date > DateTime.Now.Date;
                    }
                    // se non è data valida, NON escludo (oppure potresti voler escludere: dipende); lasciamo passare.
                    return false;
                }

            case RuleKind.MaxLength:
                {
                    int max = rule.IntParam.GetValueOrDefault(16);
                    var s = Get(rule.Field1);
                    return s != null && s.Length > max;
                }

            case RuleKind.NotPresentInPhase2:
                {
                    string dataset = rule.Phase2Dataset ?? "";
                    var val = Get(rule.Field1);
                    if (string.IsNullOrWhiteSpace(val)) return true; // assente → escludi

                    if (string.Equals(dataset, "Persone", StringComparison.OrdinalIgnoreCase))
                        return !peopleSet.Contains(val);

                    if (string.Equals(dataset, "Corsi", StringComparison.OrdinalIgnoreCase))
                        return !courseSet.Contains(val);

                    return false;
                }

            case RuleKind.PairNotPresentInPhase2:
                {
                    string dataset = rule.Phase2Dataset ?? "";
                    var a = Get(rule.Field1);
                    var b = Get(rule.Field2);
                    if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return true;

                    if (string.Equals(dataset, "Persone", StringComparison.OrdinalIgnoreCase))
                        return !peoplePairs.Contains((a, b));

                    if (string.Equals(dataset, "Corsi", StringComparison.OrdinalIgnoreCase))
                        return !coursePairs.Contains((a, b));

                    return false;
                }
        }

        return false;
    }

    // ===== Costruisci lookup da Fase 2 (legge i file/colonne scelti in Fase 2) =====
    // Persone: Col1 (es. Codice Fiscale), Col2 (es. Person Number)
    private (HashSet<string> one, HashSet<(string, string)> pairs) BuildPhase2PeopleSets()
    {
        var path = _settings.Get("Phase2PersonPath") ?? txtPersonMapPath.Text;
        var sheet = _settings.Get("Phase2PersonSheet") ?? cmbPersonSheet.SelectedItem?.ToString();
        var col1Name = _settings.Get("Phase2PersonCol1") ?? cmbPersonCol1.SelectedItem?.ToString();
        var col2Name = _settings.Get("Phase2PersonCol2") ?? cmbPersonCol2.SelectedItem?.ToString();

        return BuildSets(path, sheet, col1Name, col2Name);
    }

    // Corsi: Col1 (es. Titolo Corso), Col2 (es. Nome Corso) — o CodCorso+Titolo: dipende da come selezioni in Fase 2.
    private (HashSet<string> one, HashSet<(string, string)> pairs) BuildPhase2CourseSets()
    {
        var path = _settings.Get("Phase2CoursePath") ?? txtCourseMapPath.Text;
        var sheet = _settings.Get("Phase2CourseSheet") ?? cmbCourseSheet.SelectedItem?.ToString();
        var col1Name = _settings.Get("Phase2CourseCol1") ?? cmbCourseCol1.SelectedItem?.ToString();
        var col2Name = _settings.Get("Phase2CourseCol2") ?? cmbCourseCol2.SelectedItem?.ToString();

        return BuildSets(path, sheet, col1Name, col2Name);
    }

    // Legge l'intero foglio, individua indice colonne per nome header, estrae set singolo e set coppie
    private (HashSet<string> one, HashSet<(string, string)> pairs) BuildSets(string? path, string? sheet, string? col1Name, string? col2Name)
    {
        var one = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pairs = new HashSet<(string, string)>();

        try
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return (one, pairs);

            var (headers, rows) = ReadSheetAll(path, sheet ?? "");
            int i1 = headers.FindIndex(h => string.Equals(h, col1Name ?? "", StringComparison.OrdinalIgnoreCase));
            int i2 = headers.FindIndex(h => string.Equals(h, col2Name ?? "", StringComparison.OrdinalIgnoreCase));

            foreach (var r in rows)
            {
                if (i1 >= 0 && i1 < r.Count)
                {
                    var a = r[i1] ?? "";
                    if (!string.IsNullOrWhiteSpace(a)) one.Add(a);

                    if (i2 >= 0 && i2 < r.Count)
                    {
                        var b = r[i2] ?? "";
                        if (!string.IsNullOrWhiteSpace(a) && !string.IsNullOrWhiteSpace(b))
                            pairs.Add((a, b));
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _log.Error($"Errore nel leggere dataset Fase 2: {ex.Message}");
        }

        return (one, pairs);
    }
    // ======================
// FIX metodi mancanti
// ======================

// Popola la lista dei template salvati (Fase 3)
private void ReloadSavedTemplatesList()
{
    if (lstSavedTemplates == null) return;
    lstSavedTemplates.Items.Clear();
    foreach (var name in _templates.List())
        lstSavedTemplates.Items.Add(name);
}

// Rileva fogli e intestazioni in base al file selezionato nella Fase 1
private void TryPopulateSheetAndHeaders()
{
    cmbSheet.Items.Clear();
    cmbColumn.Items.Clear();

    var path = (txtExcel.Text ?? "").Trim();
    if (!File.Exists(path)) return;

    var sheets = _lists.GetSheets(path);
    if (sheets.Count > 0)
    {
        foreach (var s in sheets) cmbSheet.Items.Add(s);
        cmbSheet.SelectedIndex = 0;

        // ricarica intestazioni quando cambio foglio
        cmbSheet.SelectedIndexChanged -= CmbSheet_SelectedIndexChanged_LoadHeaders;
        cmbSheet.SelectedIndexChanged += CmbSheet_SelectedIndexChanged_LoadHeaders;
        LoadHeadersForCurrent();
    }
    else
    {
        // CSV/TXT o file senza fogli: carico direttamente le intestazioni "flat"
        cmbSheet.Items.Add("(non applicabile)");
        cmbSheet.SelectedIndex = 0;

        var headers = _lists.GetHeaders(path, null);
        foreach (var h in headers) cmbColumn.Items.Add(h);
        if (cmbColumn.Items.Count > 0) cmbColumn.SelectedIndex = 0;
    }
}

// Helper per collegare l'evento SelectedIndexChanged del combo "Foglio"
private void CmbSheet_SelectedIndexChanged_LoadHeaders(object? sender, EventArgs e) => LoadHeadersForCurrent();

// Legge le intestazioni del foglio selezionato e popola la combo "Colonna"
private void LoadHeadersForCurrent()
{
    cmbColumn.Items.Clear();

    var path = (txtExcel.Text ?? "").Trim();
    if (!File.Exists(path)) return;

    var sel = cmbSheet.SelectedItem?.ToString() ?? "";
    string? sheet = string.Equals(sel, "(non applicabile)", StringComparison.OrdinalIgnoreCase) ? null : sel;

    var headers = _lists.GetHeaders(path, sheet);
    foreach (var h in headers) cmbColumn.Items.Add(h);
    if (cmbColumn.Items.Count > 0) cmbColumn.SelectedIndex = 0;
}

// Suggerisce/imposta il percorso di output predefinito per la Fase 3
private void UpdateDefaultPhase3Out()
{
    var input = (txtExcel.Text ?? "").Trim();

    // Cartella base: stessa dell'input se spuntato e il file esiste; altrimenti quella scelta; fallback Desktop
    string baseDir;
    if (chkSameFolder.Checked && File.Exists(input))
        baseDir = Path.GetDirectoryName(input) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    else if (!string.IsNullOrWhiteSpace(txtOutputFolder.Text))
        baseDir = txtOutputFolder.Text.Trim();
    else
        baseDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

    // Nome base: dal file di input se disponibile; altrimenti "Output"
    string baseName = File.Exists(input) ? (Path.GetFileNameWithoutExtension(input) ?? "Output") : "Output";
    string suggested = Path.Combine(baseDir, baseName + "_Fase3.xlsx");

    // Se non c’è nulla o se è un vecchio suggerimento, imposta quello nuovo
    if (string.IsNullOrWhiteSpace(txtPhase3Out.Text)
        || txtPhase3Out.Text.EndsWith("_Fase3.xlsx", StringComparison.OrdinalIgnoreCase))
    {
        txtPhase3Out.Text = suggested;
    }
}

}
