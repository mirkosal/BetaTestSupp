using System;
using System.ComponentModel;        // LicenseManager.UsageMode
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// ==== OpenXML ====
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

// Alias per evitare ambiguità con DocumentFormat.OpenXml.Spreadsheet.Color
using WinColor = System.Drawing.Color;

public partial class MainForm : BaseMenuForm
{
    // ===== OpenXML helper =====
    private OpenXmlExcel? _ox;          // wrapper su SpreadsheetDocument
    private string? _activeSheetName;   // nome foglio selezionato

    // ===== Stato =====
    private string _excelPath = "";
    private int _currentRow = 2;         // riga 1 = header
    private int _screenColIndex = -1;    // colonna per gli screenshot (1-based)
    private int _displayColIndex = 2;    // colonna visualizzata (default B)
    private bool _suppressComboEvents = false;

    // ===== Hotkeys =====
    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID_SHOT = 1;  // ALT+S
    private const int HOTKEY_ID_NEXT = 2;  // ALT+N
    private const int HOTKEY_ID_PREV = 3;  // ALT+M
    private const uint MOD_ALT = 0x0001;

    [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    // ===== Esclusione overlay dagli screenshot (Win10/11) =====
    [DllImport("user32.dll")] private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint dwAffinity);
    private const uint WDA_EXCLUDEFROMCAPTURE = 0x11;

    // ===== Forza davvero "sempre in primo piano" =====
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(
        IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_SHOWWINDOW = 0x0040;
    private const uint SWP_NOSENDCHANGING = 0x0400;

    private void EnsureTopMost(bool on = true)
    {
        try
        {
            this.TopMost = on;
            SetWindowPos(
                this.Handle,
                on ? HWND_TOPMOST : HWND_NOTOPMOST,
                0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOSENDCHANGING
            );
        }
        catch { /* ignore */ }
    }

    // Helper item combo colonne
    private class ColItem
    {
        public string Text { get; set; } = "";
        public int Index { get; set; }   // 1-based
        public override string ToString() => Text;
    }

    public MainForm()
    {
        InitializeComponent();

        // In design-time non agganciare logica runtime
        if (DesignMode || LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;

        // Always on top robusto
        this.HandleCreated += (_, __) => EnsureTopMost(true);
        this.Shown += (_, __) => EnsureTopMost(true);
        this.Activated += (_, __) => EnsureTopMost(true);
        this.Resize += (_, __) => { if (this.WindowState == FormWindowState.Normal) EnsureTopMost(true); };

        // Posiziona in basso-destra
        this.Load += (_, __) =>
        {
            var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1200, 800);
            int margin = 16;
            this.Location = new Point(wa.Right - this.Width - margin, wa.Bottom - this.Height - margin);
            UpdateLabelLayout();
        };

        // Wrapping dinamico label
        this.SizeChanged += (_, __) => UpdateLabelLayout();

        // Wire-up eventi principali
        this.Load += MainForm_Load;
        this.FormClosed += MainForm_FormClosed;

        btnOpen.Click += BtnOpen_Click;
        btnRefresh.Click += BtnRefresh_Click;

        cmbSheets.SelectedIndexChanged += CmbSheets_SelectedIndexChanged;
        cmbRows.SelectedIndexChanged += CmbRows_SelectedIndexChanged;
        cmbDisplayColumn.SelectedIndexChanged += CmbDisplayColumn_SelectedIndexChanged;

        // Aggiungi campo
        btnAddCol.Click += BtnAddCol_Click;
        cmbNewColRel.SelectedIndexChanged += (_, __) =>
        {
            var rel = cmbNewColRel.SelectedItem?.ToString() ?? "";
            bool needAnchor = rel == "Prima di" || rel == "Dopo";
            cmbNewColAnchor.Enabled = needAnchor;
        };

        // Scelta rapida colonna screenshot
        btnScreenQuickSet.Click += BtnScreenQuickSet_Click;
    }

    // ======== Eventi form ========
    private void MainForm_Load(object? sender, EventArgs e)
    {
        RegisterHotKey(this.Handle, HOTKEY_ID_SHOT, MOD_ALT, (uint)Keys.S);
        RegisterHotKey(this.Handle, HOTKEY_ID_NEXT, MOD_ALT, (uint)Keys.N);
        RegisterHotKey(this.Handle, HOTKEY_ID_PREV, MOD_ALT, (uint)Keys.M);

        try { SetWindowDisplayAffinity(this.Handle, WDA_EXCLUDEFROMCAPTURE); } catch { /* ignore */ }

        // default per la relazione posizione nuova colonna
        cmbNewColRel.Items.Clear();
        cmbNewColRel.Items.AddRange(new object[] { "All’inizio", "Prima di", "Dopo", "Alla fine" });
        cmbNewColRel.SelectedIndex = 3; // Alla fine
        cmbNewColAnchor.Enabled = false;

        UpdateRowHeader();
        RefreshDisplayCellLabel();
    }

    private void MainForm_FormClosed(object? sender, FormClosedEventArgs e)
    {
        try { _ox?.Save(); } catch { }
        _ox?.Dispose();

        UnregisterHotKey(this.Handle, HOTKEY_ID_SHOT);
        UnregisterHotKey(this.Handle, HOTKEY_ID_NEXT);
        UnregisterHotKey(this.Handle, HOTKEY_ID_PREV);
    }

    // ======== File / Foglio ========
    private void BtnOpen_Click(object? sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", Title = "Scegli il file Excel" };
        if (ofd.ShowDialog(this) != DialogResult.OK) return;

        _excelPath = ofd.FileName;
        OpenExcel(_excelPath);
    }

    private void BtnRefresh_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_excelPath)) return;

        try
        {
            int rowBefore = _currentRow;

            _ox?.Dispose();
            _ox = OpenXmlExcel.Open(_excelPath, createIfMissing: false);

            string? sheetName = cmbSheets.SelectedItem as string;

            _suppressComboEvents = true;
            cmbSheets.Items.Clear();
            foreach (var name in _ox.GetSheetNames())
                cmbSheets.Items.Add(name);

            if (string.IsNullOrEmpty(sheetName) || !_ox.SheetExists(sheetName))
                cmbSheets.SelectedIndex = cmbSheets.Items.Count > 0 ? 0 : -1;
            else
                cmbSheets.SelectedItem = sheetName;
            _suppressComboEvents = false;

            if (cmbSheets.SelectedIndex >= 0)
            {
                _activeSheetName = cmbSheets.SelectedItem?.ToString()!;
                _ox.SetActiveSheet(_activeSheetName);

                int last = GetRowCount();
                _currentRow = Math.Max(2, Math.Min(rowBefore, last));

                PopulateHeaderCombos();
                PopulateRowsCombo();
                PopulateScreenQuickCombo();
                RefreshKeyValuesList();
                RefreshDisplayCellLabel();
                UpdateRowHeader();
                SelectRowInComboIfPresent(_currentRow);
            }

            EnsureTopMost(true);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Errore durante l'aggiornamento:\n" + ex.Message, "Errore",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void OpenExcel(string path)
    {
        try
        {
            _ox?.Dispose();
            _ox = OpenXmlExcel.Open(path, createIfMissing: true);

            _suppressComboEvents = true;
            cmbSheets.Items.Clear();
            foreach (var name in _ox.GetSheetNames())
                cmbSheets.Items.Add(name);

            if (cmbSheets.Items.Count == 0)
            {
                // crea sheet e intestazioni di default
                _activeSheetName = "Sheet1";
                _ox.CreateSheetIfMissing(_activeSheetName);
                _ox.SetActiveSheet(_activeSheetName);
                _ox.SetCellText(1, 1, "id");
                _ox.SetCellText(1, 2, "title");
                _ox.SetCellText(1, 3, "notes");
                _ox.Save();
                cmbSheets.Items.Add(_activeSheetName);
            }

            cmbSheets.SelectedIndex = 0;
            _activeSheetName = cmbSheets.SelectedItem?.ToString()!;
            _ox.SetActiveSheet(_activeSheetName);

            _suppressComboEvents = false;
            EnsureTopMost(true);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Errore apertura Excel:\n" + ex.Message, "Errore",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void CmbSheets_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressComboEvents) return;
        if (_ox == null || cmbSheets.SelectedIndex < 0) return;

        _activeSheetName = cmbSheets.SelectedItem?.ToString()!;
        _ox.SetActiveSheet(_activeSheetName);

        if (GetRowCount() == 0)
        {
            _ox.SetCellText(1, 1, "id");
            _ox.SetCellText(1, 2, "title");
            _ox.SetCellText(1, 3, "notes");
            _ox.Save();
        }

        _currentRow = 2;
        _screenColIndex = -1;
        _displayColIndex = Math.Min(Math.Max(1, _displayColIndex), Math.Max(1, GetColumnCount()));
        if (_displayColIndex < 1) _displayColIndex = 2;

        PopulateHeaderCombos();
        PopulateRowsCombo();
        PopulateScreenQuickCombo();
        RefreshKeyValuesList();
        RefreshDisplayCellLabel();
        UpdateRowHeader();
        SelectRowInComboIfPresent(_currentRow);

        EnsureTopMost(true);
    }

    // ======== Helpers intestazioni/righe/colonne ========
    private int GetColumnCount() => _ox?.GetLastColumn() ?? 0;
    private int GetRowCount() => _ox?.GetLastRow() ?? 0;

    private string[] GetHeaders()
    {
        int cols = GetColumnCount();
        if (cols <= 0 || _ox == null) return Array.Empty<string>();
        string[] hs = new string[cols];
        for (int c = 1; c <= cols; c++) hs[c - 1] = _ox.GetCellText(1, c) ?? string.Empty;
        return hs;
    }

    private static string ExcelColName(int index) // 1-based
    {
        int n = index;
        string s = "";
        while (n > 0)
        {
            n--;
            s = (char)('A' + (n % 26)) + s;
            n /= 26;
        }
        return s;
    }

    private int FindHeaderColumn(string? headerName)
    {
        if (_ox == null || string.IsNullOrEmpty(headerName)) return -1;
        int cols = GetColumnCount();
        for (int c = 1; c <= cols; c++)
        {
            var t = (_ox.GetCellText(1, c) ?? "").Trim();
            if (string.Equals(t, headerName.Trim(), StringComparison.OrdinalIgnoreCase))
                return c;
        }
        return -1;
    }

    // ======== Popola UI ========
    private void PopulateHeaderCombos()
    {
        if (_ox == null) return;
        _suppressComboEvents = true;

        cmbDisplayColumn.Items.Clear();
        cmbNewColAnchor.Items.Clear();

        int cols = GetColumnCount();
        var hs = GetHeaders();
        for (int c = 1; c <= cols; c++)
        {
            string letter = ExcelColName(c);
            string name = string.IsNullOrWhiteSpace(hs[c - 1]) ? "(senza nome)" : hs[c - 1];
            var item = new ColItem { Text = $"{letter}: {name}", Index = c };
            cmbDisplayColumn.Items.Add(item);
            cmbNewColAnchor.Items.Add(new ColItem { Text = item.Text, Index = c });
        }

        // display col default
        int wanted = (_displayColIndex >= 1 && _displayColIndex <= cols) ? _displayColIndex : (cols >= 2 ? 2 : 1);
        int idxWanted = -1;
        for (int i = 0; i < cmbDisplayColumn.Items.Count; i++)
            if (cmbDisplayColumn.Items[i] is ColItem it && it.Index == wanted) { idxWanted = i; break; }
        cmbDisplayColumn.SelectedIndex = idxWanted >= 0 ? idxWanted : (cmbDisplayColumn.Items.Count > 0 ? 0 : -1);

        cmbNewColAnchor.SelectedIndex = cmbNewColAnchor.Items.Count > 0 ? 0 : -1;

        _suppressComboEvents = false;
    }

    private void PopulateRowsCombo()
    {
        _suppressComboEvents = true;

        int previousRow = _currentRow;
        cmbRows.Items.Clear();
        if (_ox == null) { _suppressComboEvents = false; return; }

        int last = GetRowCount();
        for (int r = 2; r <= last; r++)
        {
            var txt = _ox.GetCellText(r, _displayColIndex);
            if (!string.IsNullOrWhiteSpace(txt))
                cmbRows.Items.Add(r);
        }

        int idx = -1;
        for (int i = 0; i < cmbRows.Items.Count; i++)
            if (cmbRows.Items[i] is int rr && rr == previousRow) { idx = i; break; }

        if (idx >= 0)
            cmbRows.SelectedIndex = idx;
        else if (cmbRows.Items.Count > 0)
        {
            cmbRows.SelectedIndex = 0;
            if (cmbRows.SelectedItem is int r0) _currentRow = r0;
        }

        _suppressComboEvents = false;
    }

    private void PopulateScreenQuickCombo()
    {
        _suppressComboEvents = true;
        cmbScreenQuick.Items.Clear();
        if (_ox == null) { _suppressComboEvents = false; return; }

        var hs = GetHeaders();
        for (int c = 1; c <= hs.Length; c++)
        {
            string letter = ExcelColName(c);
            string header = hs[c - 1];
            string text = $"{letter}: " + (string.IsNullOrWhiteSpace(header) ? "(senza nome)" : header);
            cmbScreenQuick.Items.Add(new ColItem { Text = text, Index = c });
        }

        if (_screenColIndex >= 1 && _screenColIndex <= hs.Length)
        {
            int idxWanted = -1;
            for (int i = 0; i < cmbScreenQuick.Items.Count; i++)
                if (cmbScreenQuick.Items[i] is ColItem it && it.Index == _screenColIndex) { idxWanted = i; break; }
            cmbScreenQuick.SelectedIndex = idxWanted >= 0 ? idxWanted : -1;
        }
        else
        {
            cmbScreenQuick.SelectedIndex = -1;
        }

        _suppressComboEvents = false;
    }

    private void SelectRowInComboIfPresent(int row)
    {
        _suppressComboEvents = true;
        int idx = -1;
        for (int i = 0; i < cmbRows.Items.Count; i++)
            if (cmbRows.Items[i] is int r && r == row) { idx = i; break; }
        if (idx >= 0) cmbRows.SelectedIndex = idx;
        _suppressComboEvents = false;
    }

    // ======== UI: lista chiave/valore ========
    private void RefreshKeyValuesList()
    {
        lstKeyValues.Items.Clear();
        if (_ox == null) return;

        int cols = GetColumnCount();
        string header(int c) => _ox.GetCellText(1, c) ?? string.Empty;
        string value(int c) => _ox.GetCellText(_currentRow, c) ?? string.Empty;

        for (int c = 1; c <= cols; c++)
            lstKeyValues.Items.Add($"{c,3}. {header(c)}: {value(c)}");

        UpdateRowHeader();
        RefreshDisplayCellLabel();
    }

    private void UpdateRowHeader()
    {
        string colVis = ExcelColName(_displayColIndex);
        lblInfo.Text =
            $"Riga: {_currentRow}  |  Righe: {GetRowCount()}  |  Foglio: {_activeSheetName ?? "(—)"}  |  Vis.: {colVis}" +
            (_screenColIndex > 0 ? $"  |  Screen: #{_screenColIndex}" : "  |  Screen: (—)");
    }

    // ======== Label contenuto testuale ========
    private void RefreshDisplayCellLabel()
    {
        if (_ox == null)
        {
            lblBValue.Text = "—";
            lblBValue.ForeColor = WinColor.Yellow;
            return;
        }

        var txt = _ox.GetCellText(_currentRow, _displayColIndex);
        if (txt != null)
        {
            string s = txt.Trim();
            lblBValue.ForeColor = WinColor.Yellow;
            lblBValue.Text = string.IsNullOrEmpty(s) ? "—" : s;
        }
        else
        {
            string coord = $"{ExcelColName(_displayColIndex)}{_currentRow}";
            lblBValue.Text = $"⚠ Contenuto non testuale in {coord}";
            lblBValue.ForeColor = WinColor.OrangeRed;
        }

        UpdateLabelLayout();
    }

    // ======== Layout dinamico label ========
    private void UpdateLabelLayout()
    {
        int margin = 24;
        int maxWidth = Math.Max(100, this.ClientSize.Width - margin * 2);
        lblBValue.MaximumSize = new Size(maxWidth, 0); // 0 = altezza auto
        lblBValue.AutoSize = true;
        lblBValue.Dock = DockStyle.Top;
        this.PerformLayout();
    }

    // ======== Aggiungi/sposta colonna + dimensioni ========
    private void BtnAddCol_Click(object? sender, EventArgs e)
    {
        if (_ox == null) return;

        var name = (txtNewColName.Text ?? "").Trim();
        if (string.IsNullOrEmpty(name))
        {
            MessageBox.Show(this, "Inserisci il nome del campo da aggiungere/spostare.", "Attenzione",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // posizione da "relazione" + "ancora"
        int cols = GetColumnCount();
        string rel = cmbNewColRel.SelectedItem?.ToString() ?? "Alla fine";
        int pos;
        if (rel == "All’inizio") pos = 1;
        else if (rel == "Alla fine") pos = cols + 1;
        else
        {
            if (cmbNewColAnchor.SelectedItem is not ColItem anchor)
            {
                MessageBox.Show(this, "Seleziona una colonna di riferimento.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            pos = (rel == "Prima di") ? anchor.Index : anchor.Index + 1;
        }

        double desiredColWidth = (double)nudNewColWidth.Value;   // 0 = non toccare
        double desiredRowHeight = (double)nudNewRowHeight.Value;  // 0 = non toccare

        int existing = FindHeaderColumn(name);
        int targetCol = pos;

        if (existing > 0)
        {
            if (!chkMoveIfExists.Checked)
            {
                MessageBox.Show(this, $"Esiste già una colonna '{name}' alla posizione {existing}. " +
                    "Abilita 'Se esiste, sposta' per riposizionarla oppure usa un altro nome.", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _ox.MoveColumn(existing, pos);
            targetCol = pos;
        }
        else
        {
            _ox.InsertColumn(pos);
            targetCol = pos;
        }

        _ox.SetCellText(1, targetCol, name);

        if (desiredColWidth > 0) _ox.SetColumnWidth(targetCol, desiredColWidth);       // approx: Excel unit
        if (desiredRowHeight > 0)
        {
            int last = Math.Max(GetRowCount(), 2);
            for (int r = 2; r <= last; r++)
                _ox.SetRowHeight(r, desiredRowHeight);                                  // points
        }

        _ox.Save();

        // Aggiorna UI
        _displayColIndex = targetCol;
        PopulateHeaderCombos();
        PopulateRowsCombo();
        PopulateScreenQuickCombo();
        RefreshKeyValuesList();
        RefreshDisplayCellLabel();
        UpdateRowHeader();
        EnsureTopMost(true);
    }

    // ======== Colonna screenshot rapida ========
    private void BtnScreenQuickSet_Click(object? sender, EventArgs e)
    {
        if (_ox == null) return;
        if (cmbScreenQuick.SelectedItem is not ColItem it)
        {
            MessageBox.Show(this, "Seleziona la colonna in cui inserire gli screenshot.", "Attenzione",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _screenColIndex = it.Index;
        MessageBox.Show(this, $"Colonna screenshot impostata su {ExcelColName(_screenColIndex)} (#{_screenColIndex}).",
            "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        UpdateRowHeader();
        EnsureTopMost(true);
    }

    // ======== Scorciatoie ========
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY)
        {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID_SHOT) { CaptureInsertAndNext(); }
            else if (id == HOTKEY_ID_NEXT) { NextRow(); }
            else if (id == HOTKEY_ID_PREV) { PrevRow(); }
        }
        base.WndProc(ref m);
    }

    private void NextRow()
    {
        _currentRow = Math.Min(GetRowCount(), _currentRow + 1);
        RefreshKeyValuesList();
        RefreshDisplayCellLabel();
        SelectRowInComboIfPresent(_currentRow);
        EnsureTopMost(true);
    }

    private void PrevRow()
    {
        _currentRow = Math.Max(2, _currentRow - 1);
        RefreshKeyValuesList();
        RefreshDisplayCellLabel();
        SelectRowInComboIfPresent(_currentRow);
        EnsureTopMost(true);
    }

    // ======== DIAGNOSTICA & SALVATAGGIO ========
    private string DumpState(string step)
    {
        string wsName = _activeSheetName ?? "(null)";
        int cols = GetColumnCount();
        int rows = GetRowCount();

        return
$@"STEP: {step}
_ox null: {_ox == null}
_ws.Name: {wsName}
_excelPath empty: {string.IsNullOrWhiteSpace(_excelPath)}
_screenColIndex: {_screenColIndex}
_currentRow: {_currentRow}
Sheet Cols: {cols}
Sheet Rows: {rows}
PrimaryScreen null: {Screen.PrimaryScreen == null}";
    }

    private bool GuardState(out string error)
    {
        error = "";
        if (_ox == null) { error = "Workbook nullo: apri un file .xlsx."; return false; }
        if (string.IsNullOrWhiteSpace(_excelPath)) { error = "Percorso file vuoto: apri un file .xlsx."; return false; }

        int maxCol = GetColumnCount();
        int maxRow = GetRowCount();

        if (_screenColIndex < 1 || _screenColIndex > Math.Max(1, maxCol))
        { error = $"Colonna screenshot fuori range (1..{maxCol})."; return false; }

        if (_currentRow < 2 || _currentRow > Math.Max(2, maxRow))
        { error = $"Riga corrente fuori range (2..{maxRow})."; return false; }

        if (Screen.PrimaryScreen == null) { error = "Nessun monitor principale disponibile."; return false; }
        return true;
    }

    private void SafeSave()
    {
        try
        {
            _ox?.Save();
        }
        catch (IOException)
        {
            // tentativo di Safe SaveAs + replace
            string dir = Path.GetDirectoryName(_excelPath!) ?? "";
            string name = Path.GetFileNameWithoutExtension(_excelPath);
            string ext = Path.GetExtension(_excelPath);
            string temp = Path.Combine(dir, $"{name}_tmp_{DateTime.Now:yyyyMMddHHmmssfff}{ext}");

            _ox!.SaveAs(temp);

            string bak = Path.Combine(dir, $"{name}.bak");
            if (File.Exists(bak)) File.Delete(bak);
            File.Replace(temp, _excelPath!, bak, true);
            if (File.Exists(temp)) File.Delete(temp);

            // ricarica dopo replace
            _ox.Dispose();
            _ox = OpenXmlExcel.Open(_excelPath, createIfMissing: false);
            _ox.SetActiveSheet(_activeSheetName ?? _ox.GetSheetNames().FirstOrDefault() ?? "Sheet1");
        }
    }

    // ======== Cattura → inserisci (drawing) → next ========
    private void CaptureInsertAndNext()
    {
        try
        {
            if (!GuardState(out string guardError))
                throw new InvalidOperationException(guardError + Environment.NewLine + DumpState("GuardState"));

            bool originallyVisible = this.Visible;
            string? tmpPng = null;

            try
            {
                if (chkHideOnCapture.Checked)
                {
                    this.Visible = false;
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(40);
                }

                var prim = Screen.PrimaryScreen;
                if (prim == null) throw new NullReferenceException("PrimaryScreen è null." + Environment.NewLine + DumpState("PrimaryScreen"));

                var bounds = prim.Bounds;
                using var bmp = new Bitmap(bounds.Width, bounds.Height);
                using (var g = Graphics.FromImage(bmp))
                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);

                // PNG temporaneo
                string picName = $"screen_r{_currentRow}_{DateTime.Now:HHmmssfff}";
                tmpPng = Path.Combine(Path.GetTempPath(), picName + ".png");
                bmp.Save(tmpPng, System.Drawing.Imaging.ImageFormat.Png);

                // Pulisci cella + immagini precedenti
                _ox!.ClearCellAndPictures(_currentRow, _screenColIndex);

                // Inserisci l’immagine come DRAWING “bordo a bordo” della cella
                int pxW = _ox.GetCellPixelWidth(_screenColIndex);
                int pxH = _ox.GetCellPixelHeight(_currentRow);
                if (pxW <= 0) pxW = 120;
                if (pxH <= 0) pxH = 90;

                _ox.InsertImageFittingCell(tmpPng, _currentRow, _screenColIndex, pxW, pxH);

                // Salvataggio robusto
                SafeSave();

                // Avanza riga e aggiorna UI
                _currentRow = Math.Min(GetRowCount(), _currentRow + 1);
                RefreshKeyValuesList();
                RefreshDisplayCellLabel();
                SelectRowInComboIfPresent(_currentRow);
            }
            finally
            {
                if (chkHideOnCapture.Checked)
                {
                    this.Visible = originallyVisible;
                    Application.DoEvents();
                    EnsureTopMost(true);
                }
            }
        }
        catch (Exception ex)
        {
            var msg = $"Errore durante cattura/salvataggio:\n{ex.Message}\n\nDettagli:\n{DumpState("Catch")}";
            try { msg += "\n\nStack:\n" + ex.StackTrace; } catch { }
            MessageBox.Show(this, msg, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ======== Handlers extra ========
    private void CmbRows_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressComboEvents) return;
        if (cmbRows.SelectedItem is int r)
        {
            _currentRow = r;
            RefreshKeyValuesList();
            RefreshDisplayCellLabel();
            EnsureTopMost(true);
        }
    }

    private void CmbDisplayColumn_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressComboEvents) return;
        if (cmbDisplayColumn.SelectedItem is ColItem it)
        {
            _displayColIndex = it.Index;
            PopulateRowsCombo();
            RefreshDisplayCellLabel();
            UpdateRowHeader();
            SelectRowInComboIfPresent(_currentRow);
            EnsureTopMost(true);
        }
    }

    // ==========================================================
    // ===============   OpenXML Helper Interno   ===============
    // ==========================================================
    private sealed class OpenXmlExcel : IDisposable
    {
        private readonly SpreadsheetDocument _doc;
        private readonly WorkbookPart _wbp;
        private WorksheetPart _activeWsp;
        private string _activeSheetName;

        private OpenXmlExcel(SpreadsheetDocument doc)
        {
            _doc = doc;
            _wbp = _doc.WorkbookPart!;
            if (_wbp.Workbook.Sheets == null) _wbp.Workbook.AppendChild(new Sheets());
            // Default active
            var first = (_wbp.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault());
            if (first == null)
            {
                _activeWsp = CreateSheet("Sheet1");
                _activeSheetName = "Sheet1";
            }
            else
            {
                _activeWsp = (WorksheetPart)_wbp.GetPartById(first.Id!);
                _activeSheetName = first.Name!;
            }
        }

        public static OpenXmlExcel Open(string path, bool createIfMissing)
        {
            if (!File.Exists(path))
            {
                if (!createIfMissing) throw new FileNotFoundException("File non trovato", path);
                using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
                {
                    var wbp = doc.AddWorkbookPart();
                    wbp.Workbook = new Workbook();
                    wbp.Workbook.AppendChild(new Sheets());
                } // using => chiusura
            }
            var sdoc = SpreadsheetDocument.Open(path, true);
            if (sdoc.WorkbookPart == null)
            {
                var wbp = sdoc.AddWorkbookPart();
                wbp.Workbook = new Workbook();
                wbp.Workbook.AppendChild(new Sheets());
            }
            return new OpenXmlExcel(sdoc);
        }

        public void Dispose()
        {
            try { _wbp?.Workbook?.Save(); } catch { /* ignore */ }
            _doc?.Dispose();
        }

        public void Save() => _wbp.Workbook.Save();

        public void SaveAs(string newPath)
        {
            using var ms = new MemoryStream();
            _doc.Clone(ms);
            File.WriteAllBytes(newPath, ms.ToArray());
        }

        public string[] GetSheetNames()
        {
            return _wbp.Workbook.Sheets!
                     .Elements<Sheet>()
                     .Select(s => s.Name?.Value ?? string.Empty)
                     .ToArray();
        }

        public bool SheetExists(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            return _wbp.Workbook.Sheets!.Elements<Sheet>()
                     .Any(s => string.Equals(s.Name?.Value ?? "", name, StringComparison.OrdinalIgnoreCase));
        }

        public void CreateSheetIfMissing(string name)
        {
            if (!SheetExists(name)) CreateSheet(name);
        }

        private WorksheetPart CreateSheet(string name)
        {
            var wsp = _wbp.AddNewPart<WorksheetPart>();
            wsp.Worksheet = new Worksheet(new SheetData());
            wsp.Worksheet.Save();

            var sheets = _wbp.Workbook.Sheets!;
            uint newId = sheets.Elements<Sheet>()
                               .Select(s => s.SheetId?.Value ?? 0U)
                               .DefaultIfEmpty(0U)
                               .Max() + 1U;

            var relId = _wbp.GetIdOfPart(wsp);
            var sheet = new Sheet() { Id = relId, SheetId = newId, Name = name };
            sheets.Append(sheet);
            _wbp.Workbook.Save();
            return wsp;
        }

        public void SetActiveSheet(string name)
        {
            var sheet = _wbp.Workbook.Sheets!.Elements<Sheet>()
                         .FirstOrDefault(s => string.Equals(s.Name?.Value ?? "", name, StringComparison.OrdinalIgnoreCase));
            if (sheet == null) throw new InvalidOperationException($"Foglio '{name}' non trovato.");

            _activeWsp = (WorksheetPart)_wbp.GetPartById(sheet.Id!);
            _activeSheetName = sheet.Name!;
        }

        private SheetData SD => _activeWsp.Worksheet.GetFirstChild<SheetData>() ?? _activeWsp.Worksheet.AppendChild(new SheetData());

        private static string RefA1(int row, int col)
        {
            string colName = "";
            int n = col;
            while (n > 0) { n--; colName = (char)('A' + (n % 26)) + colName; n /= 26; }
            return $"{colName}{row}";
        }

        private Row GetOrCreateRow(int rowIndex)
        {
            var row = SD.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = (uint)rowIndex };
                Row? after = SD.Elements<Row>().Where(r => r.RowIndex < (uint)rowIndex).OrderBy(r => r.RowIndex).LastOrDefault();
                if (after != null) SD.InsertAfter(row, after); else SD.Append(row);
            }
            return row;
        }

        private static int ColumnIndex(string cellRef)
        {
            int i = 0;
            foreach (char ch in cellRef)
            {
                if (char.IsLetter(ch))
                {
                    i = i * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
                }
                else break;
            }
            return i;
        }

        private Cell GetOrCreateCell(int rowIndex, int colIndex)
        {
            var a1 = RefA1(rowIndex, colIndex);
            var row = GetOrCreateRow(rowIndex);
            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == a1);
            if (cell == null)
            {
                Cell? after = row.Elements<Cell>()
                                 .Where(c => ColumnIndex(c.CellReference!) < colIndex)
                                 .OrderBy(c => ColumnIndex(c.CellReference!)).LastOrDefault();
                cell = new Cell() { CellReference = a1, DataType = CellValues.String };
                if (after != null) row.InsertAfter(cell, after); else row.InsertAt(cell, 0);
            }
            return cell;
        }

        public string? GetCellText(int row, int col)
        {
            var a1 = RefA1(row, col);
            var rowEl = SD.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)row);
            var cell = rowEl?.Elements<Cell>().FirstOrDefault(c => c.CellReference == a1);
            if (cell == null) return null;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sst = _wbp.SharedStringTablePart?.SharedStringTable;
                if (sst == null) return null;
                if (!int.TryParse(cell.CellValue?.InnerText ?? "0", out int idx)) return null;
                return sst.ElementAt(idx).InnerText;
            }
            return cell.CellValue?.InnerText;
        }

        public void SetCellText(int row, int col, string text)
        {
            var cell = GetOrCreateCell(row, col);
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(text ?? "");
            _activeWsp.Worksheet.Save();
        }

        public int GetLastRow()
        {
            var lastRow = SD.Elements<Row>().Where(r => r.Elements<Cell>().Any()).Select(r => (int)r.RowIndex!.Value).DefaultIfEmpty(0).Max();
            return lastRow;
        }

        public int GetLastColumn()
        {
            int max = 0;
            foreach (var r in SD.Elements<Row>())
            {
                foreach (var c in r.Elements<Cell>())
                {
                    int col = ColumnIndex(c.CellReference!);
                    if (col > max) max = col;
                }
            }
            return max;
        }

        public void InsertColumn(int atCol)
        {
            foreach (var r in SD.Elements<Row>())
            {
                foreach (var c in r.Elements<Cell>().ToList())
                {
                    int col = ColumnIndex(c.CellReference!);
                    int row = (int)r.RowIndex!.Value;
                    if (col >= atCol)
                    {
                        c.CellReference = RefA1(row, col + 1);
                    }
                }
            }
            _activeWsp.Worksheet.Save();
        }

        public void DeleteColumn(int colToDelete)
        {
            foreach (var r in SD.Elements<Row>())
            {
                foreach (var c in r.Elements<Cell>().ToList())
                {
                    int col = ColumnIndex(c.CellReference!);
                    int row = (int)r.RowIndex!.Value;
                    if (col == colToDelete)
                    {
                        c.Remove();
                    }
                    else if (col > colToDelete)
                    {
                        c.CellReference = RefA1(row, col - 1);
                    }
                }
            }
            _activeWsp.Worksheet.Save();
        }

        public void MoveColumn(int from, int to)
        {
            if (from == to) return;

            InsertColumn(to);

            int sourceIndex = from + (to <= from ? 1 : 0);
            int lastRow = GetLastRow();
            for (int r = 1; r <= lastRow; r++)
            {
                var val = GetCellText(r, sourceIndex);
                if (val != null) SetCellText(r, to, val);
                else
                {
                    var cell = GetOrCreateCell(r, to);
                    cell.CellValue = null;
                }
            }

            double? w = GetColumnWidth(sourceIndex);
            if (w.HasValue) SetColumnWidth(to, w.Value);

            DeleteColumn(sourceIndex);
            _activeWsp.Worksheet.Save();
        }

        private Columns EnsureColumns()
        {
            var cols = _activeWsp.Worksheet.GetFirstChild<Columns>();
            if (cols == null) cols = _activeWsp.Worksheet.InsertAt(new Columns(), 0);
            return cols;
        }

        public void SetColumnWidth(int col, double width)
        {
            var cols = EnsureColumns();
            var existing = cols.Elements<Column>().FirstOrDefault(c => c.Min <= (uint)col && c.Max >= (uint)col);
            if (existing == null)
            {
                var colEl = new Column()
                {
                    Min = (uint)col,
                    Max = (uint)col,
                    Width = width,
                    CustomWidth = true
                };
                cols.Append(colEl);
            }
            else
            {
                existing.Width = width;
                existing.CustomWidth = true;
            }
            _activeWsp.Worksheet.Save();
        }

        public double? GetColumnWidth(int col)
        {
            var cols = _activeWsp.Worksheet.GetFirstChild<Columns>();
            var existing = cols?.Elements<Column>().FirstOrDefault(c => c.Min <= (uint)col && c.Max >= (uint)col);
            return existing?.Width?.Value;
        }

        public void SetRowHeight(int row, double points)
        {
            var r = GetOrCreateRow(row);
            r.CustomHeight = true;
            r.Height = points;
            _activeWsp.Worksheet.Save();
        }

        public int GetCellPixelWidth(int col)
        {
            double w = GetColumnWidth(col) ?? 8.43;
            return (int)Math.Round(w * 7.0);
        }
        public int GetCellPixelHeight(int row)
        {
            var r = SD.Elements<Row>().FirstOrDefault(x => x.RowIndex == (uint)row);
            double pt = r?.Height?.Value ?? 15.0; // default excel
            return (int)Math.Round(pt * 96.0 / 72.0);
        }

        public void ClearCellAndPictures(int row, int col)
        {
            var cell = GetOrCreateCell(row, col);
            cell.CellValue = null;
            cell.DataType = null;

            var drawingsPart = _activeWsp.DrawingsPart;
            if (drawingsPart == null) { _activeWsp.Worksheet.Save(); return; }

            var wsDr = drawingsPart.WorksheetDrawing;
            if (wsDr == null) { _activeWsp.Worksheet.Save(); return; }

            var toRemove = wsDr.Elements<Xdr.TwoCellAnchor>()
                .Where(a =>
                {
                    var from = a.FromMarker;
                    var to = a.ToMarker;
                    if (from == null || to == null) return false;

                    bool okFr = int.TryParse(from.RowId?.Text, out int fr);
                    bool okFc = int.TryParse(from.ColumnId?.Text, out int fc);
                    bool okTr = int.TryParse(to.RowId?.Text, out int tr);
                    bool okTc = int.TryParse(to.ColumnId?.Text, out int tc);
                    if (!(okFr && okFc && okTr && okTc)) return false;

                    bool startsHere = (fr == row - 1 && fc == col - 1);
                    bool fullyInside = fr >= row - 1 && fc >= col - 1 && tr <= row && tc <= col;
                    return startsHere || fullyInside;
                })
                .ToList();

            foreach (var anc in toRemove)
            {
                var pic = anc.Descendants<Xdr.Picture>().FirstOrDefault();
                if (pic != null)
                {
                    var blip = pic.BlipFill?.Blip;
                    var emb = blip?.Embed?.Value;
                    if (!string.IsNullOrEmpty(emb))
                    {
                        var imgPart = drawingsPart.GetPartById(emb);
                        anc.Remove();
                        if (!drawingsPart.WorksheetDrawing.Descendants<A.Blip>().Any(b => b.Embed == emb))
                        {
                            drawingsPart.DeletePart(imgPart);
                        }
                        continue;
                    }
                }
                anc.Remove();
            }

            drawingsPart.WorksheetDrawing.Save();
            _activeWsp.Worksheet.Save();
        }

        public void InsertImageFittingCell(string imagePath, int row, int col, int pxW, int pxH)
        {
            var drawingsPart = _activeWsp.DrawingsPart ?? _activeWsp.AddNewPart<DrawingsPart>();
            if (drawingsPart.WorksheetDrawing == null)
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

            var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (var fs = File.OpenRead(imagePath))
                imagePart.FeedData(fs);

            var from = new Xdr.FromMarker(
                new Xdr.ColumnId((col - 1).ToString()),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((row - 1).ToString()),
                new Xdr.RowOffset("0"));

            var to = new Xdr.ToMarker(
                new Xdr.ColumnId(col.ToString()),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId(row.ToString()),
                new Xdr.RowOffset("0"));

            string relId = drawingsPart.GetIdOfPart(imagePart);
            var nvProps = new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = $"img_r{row}_c{col}" },
                new Xdr.NonVisualPictureDrawingProperties());

            var blipFill = new Xdr.BlipFill(
                new A.Blip() { Embed = relId },
                new A.Stretch(new A.FillRectangle()));

            var shapeProps = new Xdr.ShapeProperties(
                new A.Transform2D(
                    new A.Offset() { X = 0, Y = 0 },
                    new A.Extents() { Cx = PxToEmu(pxW), Cy = PxToEmu(pxH) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

            var pic = new Xdr.Picture(nvProps, blipFill, shapeProps);

            var anchor = new Xdr.TwoCellAnchor(
                from,
                to,
                pic,
                new Xdr.ClientData());

            drawingsPart.WorksheetDrawing.Append(anchor);

            if (_activeWsp.Worksheet.GetFirstChild<Drawing>() == null)
            {
                var drawing = new Drawing() { Id = _wbp.GetIdOfPart(drawingsPart) };
                _activeWsp.Worksheet.Append(drawing);
            }

            drawingsPart.WorksheetDrawing.Save();
            _activeWsp.Worksheet.Save();
        }

        private static long PxToEmu(int px) => (long)(px * 9525L); // 1 px = 9525 EMU (96dpi)
    }
}
