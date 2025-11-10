// CfCourseGeneratorForm.cs — versione completa SENZA namespace, con debug dock in basso e bootstrap on Load.
// - Area debug: SplitContainer (Panel1=lstLog, Panel2=textbox debug multiline, no-wrap, scrollbars).
// - Bootstrap su Load: carica fogli/intestazioni/template/regole per mostrare subito dati.
// - Logger: indirizza tutti i log anche nella debug textbox tramite DebugSink.

#nullable enable
using BetaTestSupp.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

public partial class CfCourseGeneratorForm : BaseMenuForm
{
    // ===== Servizi (creati direttamente, senza DI/Resolve) =====
    private readonly IHelpLogger? _log;
    private readonly IFileDialogService? _dialogs;
    private readonly ISettingsStore? _settings;
    private readonly ITemplateRepository? _templates;
    private readonly IListSourceReader? _lists;
    private readonly IWorkbookWriter? _writer;
    private readonly IRuleRepository? _rulesRepo;

    // ===== Stato =====
    private readonly List<string> _phase1OkRows = new();

    // ===== UI dinamica per il debug =====
    private SplitContainer? _bottomSplit;
    private TextBox? _txtDebug;

    // Storage locale per fallback JSON
    private readonly string _appData = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BetaTestSupp");

    // Property per adapter
    internal ITemplateRepository? TemplatesService => _templates;
    internal IRuleRepository? RulesService => _rulesRepo;
    internal string TemplatesFilePath => Path.Combine(_appData, "CfCourseGenerator.templates.json");
    internal string RulesFilePath => Path.Combine(_appData, "CfCourseGenerator.rules.json");

    public CfCourseGeneratorForm()
    {
        InitializeComponent(); // del Designer

        // Monta pannello debug in basso (non serve toccare il Designer)
        SetupDebugDock();

        // Collega il sink globale dei log alla textbox di debug
        DebugSink.Write = AppendDebug;

        Directory.CreateDirectory(_appData);

        // Istanziazione diretta dei servizi (se una classe non esiste, resta null ma l'UI funziona comunque)
        _log = TryCreate(() => new UiListBoxLogger(lstLog));
        _dialogs = TryCreate(() => new WinFileDialogService(this));
        _settings = TryCreate(() => new HelpFileSettingsStore("BetaTestSupp", "CfCourseGenerator.settings"));
        _templates = TryCreate(() => new HelpTemplateRepository("BetaTestSupp", "CfCourseGenerator.templates"));
        _lists = TryCreate(() => new HelpListSourceReader());
        _writer = TryCreate(() => new HelpWorkbookWriter());
        _rulesRepo = TryCreate(() => new HelpRuleRepository("BetaTestSupp", "CfCourseGenerator.rules.json"));

        WireUpUi();

        // Bootstrap su Load: popoliamo subito le combo/list per far "vedere" contenuto
        this.Load += CfCourseGeneratorForm_Load;

        _log.InfoSafe("CF Course Generator avviato.");
    }

    private static T? TryCreate<T>(Func<T> factory) where T : class
    {
        try { return factory(); } catch { return null; }
    }

    // ============================
    //  Layout: area di debug
    // ============================
    private void SetupDebugDock()
    {
        try
        {
            // Crea SplitContainer solo se non già presente
            _bottomSplit = new SplitContainer
            {
                Orientation = Orientation.Horizontal,
                Dock = DockStyle.Bottom,
                Height = 220,
                FixedPanel = FixedPanel.Panel1,
                SplitterWidth = 6
            };

            // Pannello superiore: la ListBox lstLog del Designer
            // Se la lstLog è già nella form, la stacchiamo e la rimettiamo nel Panel1
            if (lstLog != null)
            {
                // Evita doppio parent
                if (lstLog.Parent != null) lstLog.Parent.Controls.Remove(lstLog);
                lstLog.Dock = DockStyle.Fill;
                _bottomSplit.Panel1.Controls.Add(lstLog);
                _bottomSplit.Panel1MinSize = 100;
            }

            // Pannello inferiore: textbox di debug multiline che non tronca
            _txtDebug = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Dock = DockStyle.Fill,
                Font = lstLog?.Font ?? DefaultFont
            };
            _bottomSplit.Panel2.Controls.Add(_txtDebug);
            _bottomSplit.Panel2MinSize = 80;

            // Inserisci lo split in form e portalo in primo piano
            Controls.Add(_bottomSplit);
            _bottomSplit.BringToFront();

            // Teniamo margine per il contenuto sopra (tab/altro) usando Anchor/Dock del Designer
            // Non servono ulteriori modifiche al resto del layout.
        }
        catch (Exception ex)
        {
            // In caso di problemi di layout, non blocchiamo il form
            Console.WriteLine("SetupDebugDock error: " + ex.Message);
        }
    }

    private void AppendDebug(string msg)
    {
        try
        {
            if (_txtDebug == null) return;
            var line = $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}";
            _txtDebug.AppendText(line);
        }
        catch { /* ignora */ }
    }

    // ============================
    //  Bootstrap on Load
    // ============================
    private void CfCourseGeneratorForm_Load(object? sender, EventArgs e)
    {
        try
        {
            RestoreUiState();

            // Se c'è un Excel salvato, carichiamo subito fogli/headers
            if (!string.IsNullOrWhiteSpace(txtExcel?.Text) && File.Exists(txtExcel.Text))
                TrySafe(nameof(TryPopulateSheetsAndHeaders), TryPopulateSheetsAndHeaders);

            // Ricarica liste template e regole per mostrare qualcosa in Fase 3/4
            TrySafe(nameof(ReloadSavedTemplatesList), ReloadSavedTemplatesList);
            TrySafe(nameof(ReloadRulesList), ReloadRulesList);

            _log.InfoSafe("Bootstrap completato.");
        }
        catch (Exception ex)
        {
            _log.ErrorSafe("Errore in Load: " + ex.Message);
        }
    }

    private void WireUpUi()
    {
        // Fase 1
        btnBrowse?.AttachClick(() => TrySafe(nameof(BrowseExcel), BrowseExcel));
        btnOutBrowse?.AttachClick(() => TrySafe(nameof(BrowseOutputFolder), BrowseOutputFolder));
        btnGenerate?.AttachClick(() => TrySafe(nameof(RunPhase1), RunPhase1));
        if (chkSameFolder is not null)
            chkSameFolder.CheckedChanged += (_, __) => TrySafe(nameof(SyncOutputWithSource), SyncOutputWithSource);
        if (txtExcel is not null)
            txtExcel.TextChanged += (_, __) =>
            {
                TrySafe(nameof(TryPopulateSheetsAndHeaders), TryPopulateSheetsAndHeaders);
                TrySafe(nameof(UpdateDefaultPhase3Out), UpdateDefaultPhase3Out);
                TrySafe(nameof(UpdateDefaultPhase4Out), UpdateDefaultPhase4Out);
            };

        // Fase 2
        btnBrowseCourseMap?.AttachClick(() => TrySafe(nameof(BrowseCourseMap), BrowseCourseMap));
        btnBrowsePersonMap?.AttachClick(() => TrySafe(nameof(BrowsePersonMap), BrowsePersonMap));
        rdoCourseSame?.AttachChecked(() => TrySafe(nameof(RefreshCourseControls), RefreshCourseControls));
        rdoCourseOther?.AttachChecked(() => TrySafe(nameof(RefreshCourseControls), RefreshCourseControls));
        rdoPersonSame?.AttachChecked(() => TrySafe(nameof(RefreshPersonControls), RefreshPersonControls));
        rdoPersonOther?.AttachChecked(() => TrySafe(nameof(RefreshPersonControls), RefreshPersonControls));

        // Fase 3
        btnAddLiteral?.AttachClick(() => TrySafe(nameof(AddLiteralToField), AddLiteralToField));
        btnClearField?.AttachClick(() => TrySafe(nameof(ClearFieldSelection), ClearFieldSelection));
        btnAddFieldToTemplate?.AttachClick(() => TrySafe(nameof(AddFieldToTemplate), AddFieldToTemplate));
        btnRemoveField?.AttachClick(() => TrySafe(nameof(RemoveSelectedField), RemoveSelectedField));
        btnRenameField?.AttachClick(() => TrySafe(nameof(RenameSelectedField), RenameSelectedField));
        btnSaveTemplate?.AttachClick(() => TrySafe(nameof(SaveTemplate), SaveTemplate));
        btnLoadTemplate?.AttachClick(() => TrySafe(nameof(LoadSelectedTemplate), LoadSelectedTemplate));
        btnDeleteTemplate?.AttachClick(() => TrySafe(nameof(DeleteSelectedTemplate), DeleteSelectedTemplate));
        btnGeneratePhase3?.AttachClick(() => TrySafe(nameof(RunPhase3), RunPhase3));
        btnBrowsePhase3Out?.AttachClick(() => TrySafe(nameof(BrowsePhase3Out), BrowsePhase3Out));
        btnPreviewPhase3?.AttachClick(() => TrySafe(nameof(PreviewPhase3), PreviewPhase3));
        btnOpenPhase3Folder?.AttachClick(() => TrySafe(nameof(OpenPhase3Folder), OpenPhase3Folder));

        // Fase 4
        btnBrowsePhase4Out?.AttachClick(() => TrySafe(nameof(BrowsePhase4Out), BrowsePhase4Out));
        btnOpenPhase4Folder?.AttachClick(() => TrySafe(nameof(OpenPhase4Folder), OpenPhase4Folder));
        btnReloadFase3Headers?.AttachClick(() => TrySafe(nameof(ReloadFase3Headers), ReloadFase3Headers));
        btnAddRule?.AttachClick(() => TrySafe(nameof(AddRule), AddRule));
        btnEditRule?.AttachClick(() => TrySafe(nameof(EditSelectedRule), EditSelectedRule));
        btnDeleteRule?.AttachClick(() => TrySafe(nameof(DeleteSelectedRule), DeleteSelectedRule));
        btnApplyPhase4?.AttachClick(() => TrySafe(nameof(ApplyPhase4), ApplyPhase4));
    }

    private void RestoreUiState()
    {
        txtCourseMapPath?.SetText(_settings?.Get("Phase2CourseMapPath") ?? string.Empty);
        txtPersonMapPath?.SetText(_settings?.Get("Phase2PersonMapPath") ?? string.Empty);
        txtExcel?.SetText(_settings?.Get("Phase1ExcelPath") ?? string.Empty);
        txtOutputFolder?.SetText(_settings?.Get("Phase1OutDir") ?? string.Empty);
        txtPhase3Out?.SetText(_settings?.Get("Phase3OutDir") ?? string.Empty);
        txtPhase4OutDir?.SetText(_settings?.Get("Phase4OutDir") ?? string.Empty);

        RefreshCourseControls();
        RefreshPersonControls();

        TrySafe(nameof(ReloadSavedTemplatesList), ReloadSavedTemplatesList);
        TrySafe(nameof(ReloadRulesList), ReloadRulesList);
    }

    // ============================
    //  Helper robusti
    // ============================
    private void TrySafe(string op, Action act)
    {
        try { act(); }
        catch (Exception ex) { _log.ErrorSafe($"[{op}] {ex.GetType().Name}: {ex.Message}"); }
    }

    private static bool IsFile(string? p) => !string.IsNullOrWhiteSpace(p) && File.Exists(p);
    private static bool IsDir(string? p) => !string.IsNullOrWhiteSpace(p) && Directory.Exists(p);

    // ============================
    //  FASE 1
    // ============================
    private void BrowseExcel()
    {
        var p = _dialogs.OpenFileFlex("Scegli file Excel", "Excel (*.xlsx)|*.xlsx|Tutti i file (*.*)|*.*", this);
        if (!string.IsNullOrWhiteSpace(p))
        {
            txtExcel?.SetText(p);
            _settings?.Set("Phase1ExcelPath", p);
            _log.InfoSafe($"Excel sorgente: {Path.GetFileName(p)}");
        }
    }

    private void BrowseOutputFolder()
    {
        var d = _dialogs.SelectFolderFlex("Scegli cartella di output Fase 1", this);
        if (!string.IsNullOrWhiteSpace(d))
        {
            txtOutputFolder?.SetText(d);
            _settings?.Set("Phase1OutDir", d);
            _log.InfoSafe($"Cartella output Fase 1: {d}");
        }
    }

    private void SyncOutputWithSource()
    {
        if (chkSameFolder is null || txtExcel is null || txtOutputFolder is null) return;
        if (chkSameFolder.Checked && IsFile(txtExcel.Text))
        {
            var dir = Path.GetDirectoryName(txtExcel.Text) ?? string.Empty;
            txtOutputFolder.Text = dir;
            _settings?.Set("Phase1OutDir", dir);
        }
    }

    private void TryPopulateSheetsAndHeaders()
    {
        if (txtExcel is null || !IsFile(txtExcel.Text))
        {
            _log.WarnSafe("Fase 1: seleziona un file Excel valido.");
            return;
        }

        var sheets = _lists.ListSheetsFlex(txtExcel.Text);
        if (cmbSheet is not null)
        {
            cmbSheet.BeginUpdate();
            cmbSheet.Items.Clear();
            foreach (var s in sheets) cmbSheet.Items.Add(s);
            cmbSheet.EndUpdate();
            if (cmbSheet.Items.Count > 0) cmbSheet.SelectedIndex = 0;
        }

        var selSheet = cmbSheet?.SelectedItem?.ToString() ?? string.Empty;
        var headers = _lists.ListHeadersFlex(txtExcel.Text, selSheet);
        if (cmbColumn is not null)
        {
            cmbColumn.BeginUpdate();
            cmbColumn.Items.Clear();
            foreach (var h in headers) cmbColumn.Items.Add(h);
            cmbColumn.EndUpdate();
            if (cmbColumn.Items.Count > 0) cmbColumn.SelectedIndex = 0;
        }

        _log.InfoSafe("Fase 1: fogli e intestazioni caricati.");
    }

    private void RunPhase1()
    {
        if (pgbPhase1 is not null) pgbPhase1.Value = 0;
        _phase1OkRows.Clear();

        if (txtExcel is null || !IsFile(txtExcel.Text)) { _log.WarnSafe("Seleziona un file Excel valido."); return; }
        if (txtOutputFolder is null || string.IsNullOrWhiteSpace(txtOutputFolder.Text)) { _log.WarnSafe("Imposta la cartella di output."); return; }

        var sheet = cmbSheet?.SelectedItem?.ToString() ?? string.Empty;
        var column = cmbColumn?.SelectedItem?.ToString() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(column)) { _log.WarnSafe("Seleziona foglio e colonna."); return; }

        var rows = _lists.ReadOkRowsFlex(txtExcel.Text, sheet, column);
        _phase1OkRows.AddRange(rows);

        if (pgbPhase1 is not null) pgbPhase1.Value = 100;
        _log.InfoSafe($"Fase 1 completata. Righe valide: {_phase1OkRows.Count}");
    }

    // ============================
    //  FASE 2
    // ============================
    private void RefreshCourseControls()
    {
        bool enableOther = rdoCourseOther?.Checked ?? false;
        if (txtCourseMapPath is not null) txtCourseMapPath.Enabled = enableOther;
        if (btnBrowseCourseMap is not null) btnBrowseCourseMap.Enabled = enableOther;
        cmbCourseSheet?.Enable();
        cmbCourseCol1?.Enable();
        cmbCourseCol2?.Enable();
    }

    private void RefreshPersonControls()
    {
        bool enableOther = rdoPersonOther?.Checked ?? false;
        if (txtPersonMapPath is not null) txtPersonMapPath.Enabled = enableOther;
        if (btnBrowsePersonMap is not null) btnBrowsePersonMap.Enabled = enableOther;
        cmbPersonSheet?.Enable();
        cmbPersonCol1?.Enable();
        cmbPersonCol2?.Enable();
    }

    private void BrowseCourseMap()
    {
        var p = _dialogs.OpenFileFlex("Scegli CourseMap", "Excel (*.xlsx)|*.xlsx|Tutti i file (*.*)|*.*", this);
        if (!string.IsNullOrWhiteSpace(p))
        {
            txtCourseMapPath?.SetText(p);
            _settings?.Set("Phase2CourseMapPath", p);
            _log.InfoSafe($"CourseMap: {Path.GetFileName(p)}");
        }
    }

    private void BrowsePersonMap()
    {
        var p = _dialogs.OpenFileFlex("Scegli PersonMap", "Excel (*.xlsx)|*.xlsx|Tutti i file (*.*)|*.*", this);
        if (!string.IsNullOrWhiteSpace(p))
        {
            txtPersonMapPath?.SetText(p);
            _settings?.Set("Phase2PersonMapPath", p);
            _log.InfoSafe($"PersonMap: {Path.GetFileName(p)}");
        }
    }

    // ============================
    //  FASE 3
    // ============================
    private void ReloadSavedTemplatesList()
    {
        var all = this.TemplateListNamesFlex()
                 .Where(s => !string.IsNullOrWhiteSpace(s))
                 .Select(s => s!)
                 .OrderBy(s => s)
                 .ToList();

        if (lstSavedTemplates is not null)
        {
            lstSavedTemplates.BeginUpdate();
            lstSavedTemplates.Items.Clear();
            foreach (var n in all) lstSavedTemplates.Items.Add(n);
            lstSavedTemplates.EndUpdate();
        }

        _log.InfoSafe($"Template salvati: {all.Count}");
    }

    private void AddLiteralToField()
    {
        var literal = txtLiteral?.Text?.Trim();
        if (string.IsNullOrEmpty(literal)) return;
        var chip = new Label { Text = literal, AutoSize = true, Padding = new Padding(6), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(3) };
        pnlFieldChips?.Controls.Add(chip);
        txtLiteral!.Text = string.Empty;
    }

    private void ClearFieldSelection() => pnlFieldChips?.Controls.Clear();

    private IEnumerable<string> CurrentFieldTokens()
    {
        if (pnlFieldChips is null) return Enumerable.Empty<string>();
        var tokens = new List<string>();
        foreach (Control c in pnlFieldChips.Controls)
            if (!string.IsNullOrWhiteSpace(c.Text)) tokens.Add(c.Text);
        return tokens;
    }

    private void AddFieldToTemplate()
    {
        var composed = string.Join("", CurrentFieldTokens());
        if (string.IsNullOrWhiteSpace(composed)) return;
        lstTemplateFields?.Items.Add(composed);
        ClearFieldSelection();
    }

    private void RemoveSelectedField()
    {
        if (lstTemplateFields is null) return;
        var idx = lstTemplateFields.SelectedIndex;
        if (idx >= 0) lstTemplateFields.Items.RemoveAt(idx);
    }

    private void RenameSelectedField()
    {
        if (lstTemplateFields is null) return;
        var idx = lstTemplateFields.SelectedIndex;
        if (idx < 0) return;

        var old = lstTemplateFields.Items[idx]?.ToString() ?? "";
        var input = _dialogs.PromptFlex(this, "Rinomina campo", old) ?? old;
        if (!string.IsNullOrWhiteSpace(input))
            lstTemplateFields.Items[idx] = input;
    }

    private void SaveTemplate()
    {
        var name = txtTemplateName?.Text?.Trim();
        if (string.IsNullOrWhiteSpace(name)) { _log.WarnSafe("Inserisci un nome template."); return; }

        var fields = lstTemplateFields?.Items.Cast<object?>()
                       .Select(x => x?.ToString())
                       .Where(s => !string.IsNullOrWhiteSpace(s))
                       .Select(s => s!)
                       .ToList() ?? new List<string>();

        this.TemplateSaveFlex(name!, fields);
        _log.InfoSafe($"Template '{name}' salvato ({fields.Count} campi).");
        ReloadSavedTemplatesList();
    }

    private void LoadSelectedTemplate()
    {
        var sel = lstSavedTemplates?.SelectedItem?.ToString();
        if (string.IsNullOrWhiteSpace(sel)) return;

        var fields = this.TemplateLoadFlex(sel!);
        if (lstTemplateFields is not null)
        {
            lstTemplateFields.BeginUpdate();
            lstTemplateFields.Items.Clear();
            foreach (var f in fields) lstTemplateFields.Items.Add(f);
            lstTemplateFields.EndUpdate();
        }

        txtTemplateName?.SetText(sel!);
        _log.InfoSafe($"Template '{sel}' caricato ({fields.Length} campi).");
    }

    private void DeleteSelectedTemplate()
    {
        var sel = lstSavedTemplates?.SelectedItem?.ToString();
        if (string.IsNullOrWhiteSpace(sel)) return;

        if (_dialogs.ConfirmFlex(this, $"Eliminare il template '{sel}'?") == true)
        {
            this.TemplateDeleteFlex(sel!);
            _log.InfoSafe($"Template '{sel}' eliminato.");
            ReloadSavedTemplatesList();
        }
    }

    private void BrowsePhase3Out()
    {
        var d = _dialogs.SelectFolderFlex("Scegli cartella output Fase 3", this);
        if (!string.IsNullOrWhiteSpace(d))
        {
            txtPhase3Out?.SetText(d);
            _settings?.Set("Phase3OutDir", d);
            _log.InfoSafe($"Cartella output Fase 3: {d}");
        }
    }

    private void UpdateDefaultPhase3Out()
    {
        var dir = _settings?.Get("Phase3OutDir");
        if (!string.IsNullOrWhiteSpace(dir)) txtPhase3Out?.SetText(dir);
    }

    private void RunPhase3()
    {
        if (pgbPhase3 is not null) pgbPhase3.Value = 0;

        var outDir = txtPhase3Out?.Text;
        if (string.IsNullOrWhiteSpace(outDir)) { _log.WarnSafe("Imposta la cartella di output Fase 3."); return; }
        if (!IsDir(outDir)) Directory.CreateDirectory(outDir);

        var preview = new List<Dictionary<string, string>>();
        foreach (var r in _phase1OkRows.Take(50))
            preview.Add(new Dictionary<string, string> { ["ROW"] = r });

        if (dgvPhase3Preview is not null)
        {
            dgvPhase3Preview.AutoGenerateColumns = true;
            dgvPhase3Preview.DataSource = preview.Select(d => d.ToDictionary(kv => kv.Key, kv => (object)kv.Value)).ToList();
        }

        if (pgbPhase3 is not null) pgbPhase3.Value = 100;
        _log.InfoSafe("Fase 3 completata (anteprima generata).");
    }

    private void PreviewPhase3() => _log.InfoSafe("Anteprima Fase 3 aggiornata.");
    private void OpenPhase3Folder()
    {
        var d = txtPhase3Out?.Text;
        if (string.IsNullOrWhiteSpace(d) || !Directory.Exists(d)) { _log.WarnSafe("Cartella Fase 3 non valida."); return; }
        try { System.Diagnostics.Process.Start("explorer.exe", d); }
        catch (Exception ex) { _log.ErrorSafe($"Errore apertura cartella: {ex.Message}"); }
    }

    // ============================
    //  FASE 4
    // ============================
    private void ReloadRulesList()
    {
        var rules = this.RulesGetAllFlex();
        _log.InfoSafe($"Fase 4: regole caricate = {rules.Length}");
    }

    private void UpdateDefaultPhase4Out()
    {
        var d = _settings?.Get("Phase4OutDir");
        if (!string.IsNullOrWhiteSpace(d)) txtPhase4OutDir?.SetText(d);
    }

    private void BrowsePhase4Out()
    {
        var d = _dialogs.SelectFolderFlex("Scegli cartella output Fase 4", this);
        if (!string.IsNullOrWhiteSpace(d))
        {
            txtPhase4OutDir?.SetText(d);
            _settings?.Set("Phase4OutDir", d);
            _log.InfoSafe($"Cartella output Fase 4: {d}");
        }
    }

    private void OpenPhase4Folder()
    {
        var d = txtPhase4OutDir?.Text;
        if (string.IsNullOrWhiteSpace(d) || !Directory.Exists(d)) { _log.WarnSafe("Cartella Fase 4 non valida."); return; }
        try { System.Diagnostics.Process.Start("explorer.exe", d); }
        catch (Exception ex) { _log.ErrorSafe($"Errore apertura cartella: {ex.Message}"); }
    }

    private void ReloadFase3Headers()
    {
        TrySafe(nameof(TryPopulateSheetsAndHeaders), TryPopulateSheetsAndHeaders);
        _log.InfoSafe("Headers Fase 3 ricaricati.");
    }

    private void AddRule() => _log.InfoSafe("Aggiungi regola (placeholder).");
    private void EditSelectedRule() => _log.InfoSafe("Modifica regola selezionata (placeholder).");
    private void DeleteSelectedRule() => _log.InfoSafe("Elimina regola selezionata (placeholder).");

    private void ApplyPhase4()
    {
        var dir = txtPhase4OutDir?.Text;
        if (string.IsNullOrWhiteSpace(dir)) { _log.WarnSafe("Imposta la cartella di output Fase 4."); return; }
        if (!IsDir(dir)) Directory.CreateDirectory(dir);
        _log.InfoSafe("Fase 4 applicata (placeholder).");
    }
}

// ============================================================
// ===============  LOG SINK & EXTENSIONS  ====================
// ============================================================

internal static class DebugSink
{
    // Viene impostato nel costruttore del form
    public static Action<string>? Write;
}

internal static class LoggerExtensions
{
    public static void InfoSafe(this IHelpLogger? log, string msg) => log.CallOrFallback("Info", msg, "[INFO] " + msg);
    public static void ErrorSafe(this IHelpLogger? log, string msg) => log.CallOrFallback("Error", msg, "[ERROR] " + msg);
    public static void DebugSafe(this IHelpLogger? log, string msg) => log.CallOrFallback("Debug", msg, "[DEBUG] " + msg);
    public static void WarnSafe(this IHelpLogger? log, string msg)
    {
        if (log is null)
        {
            Console.WriteLine("[WARN] " + msg);
            DebugSink.Write?.Invoke(msg);
            return;
        }
        var m = log.GetType().GetMethod("Warn", new[] { typeof(string) });
        if (m != null) { m.Invoke(log, new object[] { msg }); DebugSink.Write?.Invoke(msg); return; }
        log.CallOrFallback("Info", "[WARN] " + msg, "[WARN] " + msg);
    }

    private static void CallOrFallback(this IHelpLogger? log, string method, string msg, string fallback)
    {
        DebugSink.Write?.Invoke(msg);
        if (log is null) { Console.WriteLine(fallback); return; }
        var m = log.GetType().GetMethod(method, new[] { typeof(string) });
        if (m != null) m.Invoke(log, new object[] { msg });
        else Console.WriteLine(fallback);
    }
}

internal static class FileDialogServiceExtensions
{
    public static string OpenFileFlex(this IFileDialogService? svc, string title, string filter, IWin32Window? owner = null)
    {
        // Prova i metodi dell’interfaccia, ma cattura tutte le eccezioni e fai fallback al dialog nativo
        if (svc != null)
        {
            try
            {
                var t = svc.GetType();
                var m = t.GetMethod("OpenFile", new[] { typeof(string), typeof(string) });
                if (m != null)
                {
                    var r = m.Invoke(svc, new object[] { title, filter }) as string;
                    if (!string.IsNullOrEmpty(r)) return r;
                }

                m = t.GetMethod("OpenFile", new[] { typeof(string) });
                if (m != null)
                {
                    var r = m.Invoke(svc, new object[] { filter }) as string;
                    if (!string.IsNullOrEmpty(r)) return r;
                }

                m = t.GetMethod("OpenFile", Type.EmptyTypes);
                if (m != null)
                {
                    var r = m.Invoke(svc, null) as string;
                    if (!string.IsNullOrEmpty(r)) return r;
                }
            }
            catch
            {
                // qualsiasi eccezione (incl. TargetInvocationException) -> fallback nativo
            }
        }

        using var dlg = new OpenFileDialog
        {
            Title = title,
            Filter = filter,
            CheckFileExists = true
        };
        return dlg.ShowDialog(owner) == DialogResult.OK ? dlg.FileName : string.Empty;
    }


    public static string SelectFolderFlex(this IFileDialogService? svc, string title, IWin32Window? owner = null)
    {
        if (svc != null)
        {
            try
            {
                var t = svc.GetType();
                var m = t.GetMethod("SelectFolder", new[] { typeof(string) });
                if (m != null)
                {
                    var r = m.Invoke(svc, new object[] { title }) as string;
                    if (!string.IsNullOrEmpty(r)) return r;
                }

                m = t.GetMethod("SelectFolder", Type.EmptyTypes);
                if (m != null)
                {
                    var r = m.Invoke(svc, null) as string;
                    if (!string.IsNullOrEmpty(r)) return r;
                }
            }
            catch
            {
                // fallback nativo
            }
        }

        using var dlg = new FolderBrowserDialog
        {
            Description = title,
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };
        return dlg.ShowDialog(owner) == DialogResult.OK ? dlg.SelectedPath : string.Empty;
    }

    public static string? PromptFlex(this IFileDialogService? svc, IWin32Window? owner, string title, string defaultText = "")
    {
        if (svc != null)
        {
            var t = svc.GetType();
            var m = t.GetMethod("Prompt", new[] { typeof(string), typeof(string) });
            if (m != null) return m.Invoke(svc, new object[] { title, defaultText }) as string;
            m = t.GetMethod("Prompt", new[] { typeof(string) });
            if (m != null) return m.Invoke(svc, new object[] { title }) as string;
        }
        using var f = new Form { Text = title, StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false, Width = 420, Height = 150 };
        var tb = new TextBox { Dock = DockStyle.Top, Text = defaultText, Margin = new Padding(8) };
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Dock = DockStyle.Right, Width = 80 };
        var cancel = new Button { Text = "Annulla", DialogResult = DialogResult.Cancel, Dock = DockStyle.Right, Width = 80 };
        var panel = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(8) };
        panel.Controls.Add(cancel); panel.Controls.Add(ok);
        f.Controls.Add(tb); f.Controls.Add(panel);
        return f.ShowDialog(owner) == DialogResult.OK ? tb.Text : null;
    }

    public static bool? ConfirmFlex(this IFileDialogService? svc, IWin32Window? owner, string message, string title = "Conferma")
    {
        if (svc != null)
        {
            var t = svc.GetType();
            var m = t.GetMethod("Confirm", new[] { typeof(string), typeof(string) });
            if (m != null) return (bool?)m.Invoke(svc, new object[] { message, title });
            m = t.GetMethod("Confirm", new[] { typeof(string) });
            if (m != null) return (bool?)m.Invoke(svc, new object[] { message });
        }
        var res = MessageBox.Show(owner, message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        return res == DialogResult.Yes;
    }
}

internal static class ListSourceReaderExtensions
{
    public static IEnumerable<string> ListSheetsFlex(this IListSourceReader? svc, string file)
    {
        if (svc == null) return Enumerable.Empty<string>();
        var t = svc.GetType();
        var m = t.GetMethod("ListSheets", new[] { typeof(string) }) ??
                t.GetMethod("ListWorksheets", new[] { typeof(string) }) ??
                t.GetMethod("GetSheets", new[] { typeof(string) });
        var res = m?.Invoke(svc, new object[] { file }) as IEnumerable<string>;
        return res ?? Enumerable.Empty<string>();
    }

    public static IEnumerable<string> ListHeadersFlex(this IListSourceReader? svc, string file, string sheet)
    {
        if (svc == null) return Enumerable.Empty<string>();
        var t = svc.GetType();
        var m = t.GetMethod("ListHeaders", new[] { typeof(string), typeof(string) }) ??
                t.GetMethod("GetHeaders", new[] { typeof(string), typeof(string) }) ??
                t.GetMethod("ReadHeaders", new[] { typeof(string), typeof(string) });
        var res = m?.Invoke(svc, new object[] { file, sheet }) as IEnumerable<string>;
        return res ?? Enumerable.Empty<string>();
    }

    public static IEnumerable<string> ReadOkRowsFlex(this IListSourceReader? svc, string file, string sheet, string column)
    {
        if (svc != null)
        {
            var t = svc.GetType();
            foreach (var name in new[] { "ReadOkRows", "ReadColumnValues", "GetColumnValues", "EnumerateColumn", "ReadNonEmpty" })
            {
                var m = t.GetMethod(name, new[] { typeof(string), typeof(string), typeof(string) });
                if (m != null && typeof(System.Collections.IEnumerable).IsAssignableFrom(m.ReturnType))
                {
                    try
                    {
                        var res = (m.Invoke(svc, new object[] { file, sheet, column }) as System.Collections.IEnumerable)
                                  ?.Cast<object?>().Select(o => o?.ToString() ?? "").Where(s => !string.IsNullOrWhiteSpace(s)) ?? Enumerable.Empty<string>();
                        return res!;
                    }
                    catch { /* fallback */ }
                }
            }
        }
        return Enumerable.Empty<string>();
    }
}

internal static class TemplateRepositoryFlex
{
    public static IEnumerable<string?> TemplateListNamesFlex(this CfCourseGeneratorForm self)
    {
        var svc = self.TemplatesService;
        if (svc != null)
        {
            var t = svc.GetType();
            foreach (var name in new[] { "ListTemplateNames", "ListNames", "GetNames", "AllNames" })
            {
                var m = t.GetMethod(name, Type.EmptyTypes);
                if (m != null)
                {
                    var res = m.Invoke(svc, null) as System.Collections.IEnumerable;
                    if (res != null) return res.Cast<object?>().Select(o => o?.ToString());
                }
            }
        }
        var map = self.ReadTemplatesJson();
        return map.Keys;
    }

    public static void TemplateSaveFlex(this CfCourseGeneratorForm self, string name, List<string> fields)
    {
        var svc = self.TemplatesService;
        if (svc != null)
        {
            var t = svc.GetType();
            var m = t.GetMethod("SaveTemplate", new[] { typeof(string), typeof(IEnumerable<string>) }) ??
                    t.GetMethod("Save", new[] { typeof(string), typeof(IEnumerable<string>) });
            if (m != null) { m.Invoke(svc, new object[] { name, fields }); return; }
        }
        var map = self.ReadTemplatesJson();
        map[name] = fields;
        self.WriteTemplatesJson(map);
    }

    public static string[] TemplateLoadFlex(this CfCourseGeneratorForm self, string name)
    {
        var svc = self.TemplatesService;
        if (svc != null)
        {
            var t = svc.GetType();
            var m = t.GetMethod("LoadTemplate", new[] { typeof(string) }) ??
                    t.GetMethod("Load", new[] { typeof(string) });
            if (m != null)
            {
                var res = m.Invoke(svc, new object[] { name }) as System.Collections.IEnumerable;
                if (res != null) return res.Cast<object?>().Select(o => o?.ToString() ?? "").Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
            }
        }
        var map = self.ReadTemplatesJson();
        return map.TryGetValue(name, out var fields) ? fields.ToArray() : Array.Empty<string>();
    }

    public static void TemplateDeleteFlex(this CfCourseGeneratorForm self, string name)
    {
        var svc = self.TemplatesService;
        if (svc != null)
        {
            var t = svc.GetType();
            var m = t.GetMethod("DeleteTemplate", new[] { typeof(string) }) ??
                    t.GetMethod("Delete", new[] { typeof(string) });
            if (m != null) { m.Invoke(svc, new object[] { name }); return; }
        }
        var map = self.ReadTemplatesJson();
        if (map.Remove(name)) self.WriteTemplatesJson(map);
    }

    private static Dictionary<string, List<string>> ReadTemplatesJson(this CfCourseGeneratorForm self)
    {
        try
        {
            var path = self.TemplatesFilePath;
            if (!File.Exists(path)) return new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<Dictionary<string, List<string>>>(json) ?? new();
        }
        catch { return new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase); }
    }

    private static void WriteTemplatesJson(this CfCourseGeneratorForm self, Dictionary<string, List<string>> map)
    {
        try
        {
            var json = JsonSerializer.Serialize(map, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(self.TemplatesFilePath, json);
        }
        catch { /* ignore */ }
    }
}

internal static class RuleRepositoryFlex
{
    public static RuleModelFlexible[] RulesGetAllFlex(this CfCourseGeneratorForm self)
    {
        var svc = self.RulesService;
        if (svc != null)
        {
            var t = svc.GetType();
            var m = t.GetMethod("GetAllRules", Type.EmptyTypes) ??
                    t.GetMethod("GetAll", Type.EmptyTypes) ??
                    t.GetMethod("List", Type.EmptyTypes);
            if (m != null)
            {
                var res = m.Invoke(svc, null) as System.Collections.IEnumerable;
                if (res != null)
                {
                    return res.Cast<object?>().Select(ToFlexible).Where(x => x != null).Cast<RuleModelFlexible>().ToArray();
                }
            }
        }
        return self.ReadRulesJson();
    }

    private static RuleModelFlexible? ToFlexible(object? obj)
    {
        if (obj == null) return null;
        var t = obj.GetType();
        string name = t.GetProperty("Name")?.GetValue(obj)?.ToString()
                      ?? t.GetProperty("Title")?.GetValue(obj)?.ToString() ?? "Rule";
        string type = t.GetProperty("Type")?.GetValue(obj)?.ToString() ?? "";
        string expr = t.GetProperty("Expression")?.GetValue(obj)?.ToString() ?? "";
        return new RuleModelFlexible { Name = name, Type = type, Expression = expr };
    }

    private static RuleModelFlexible[] ReadRulesJson(this CfCourseGeneratorForm self)
    {
        try
        {
            var path = self.RulesFilePath;
            if (!File.Exists(path)) return Array.Empty<RuleModelFlexible>();
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<RuleModelFlexible[]>(json) ?? Array.Empty<RuleModelFlexible>();
        }
        catch { return Array.Empty<RuleModelFlexible>(); }
    }
}

internal class RuleModelFlexible
{
    public string Name { get; set; } = "";
    public string Type { get; set; } = "";
    public string Expression { get; set; } = "";
}

internal static class UiExtensions
{
    public static void SetText(this TextBox? tb, string value)
    {
        if (tb is null) return;
        tb.Text = value ?? string.Empty;
    }

    public static void AttachClick(this Button? b, Action handler)
    {
        if (b is null) return;
        b.Click += (_, __) => handler();
    }

    public static void AttachChecked(this RadioButton? rb, Action handler)
    {
        if (rb is null) return;
        rb.CheckedChanged += (_, __) => handler();
    }

    public static void Enable(this Control? c)
    {
        if (c is null) return;
        c.Enabled = true;
    }
}
