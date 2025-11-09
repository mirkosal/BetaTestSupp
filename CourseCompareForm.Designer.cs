using System.Drawing;
using System.Windows.Forms;

public partial class CourseCompareForm : BaseMenuForm
{
    private TableLayoutPanel root;

    private Label lblTitle;

    private Label lblFileA;
    private TextBox txtFileA;
    private Button btnBrowseA;
    private Label lblColsA;
    private Label lblA1;
    private Label lblA2;
    private ListBox lstACol1;
    private ListBox lstACol2;

    private Label lblFileB;
    private TextBox txtFileB;
    private Button btnBrowseB;
    private Label lblColsB;
    private Label lblB1;
    private Label lblB2;
    private ListBox lstBCol1;
    private ListBox lstBCol2;

    private GroupBox grpOutputFrom;
    private RadioButton rdoOutFromA;
    private RadioButton rdoOutFromB;

    private FlowLayoutPanel pnlActions;
    private Button btnLoadHeaders;
    private Button btnRun;
    private Label lblStatus;

    private void InitializeComponent()
    {
        Text = "Valutazione corsi — Confronto su due colonne";
        Width = 1000;
        Height = 640;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(24, 24, 28);
        ForeColor = Color.Lime;

        root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 6,
            RowCount = 10,
            BackColor = Color.FromArgb(24, 24, 28),
            ForeColor = Color.Lime,
        };
        // Colonne: label, textbox (stretch), browse, space, lists (320px), lists (320px)
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 12));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320));
        for (int r = 0; r < root.RowCount; r++)
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        ContentHost.Controls.Add(root);

        lblTitle = new Label
        {
            Text = "Confronta File A e File B usando due colonne ciascuno (chiave composta):" +
                   " se non c’è match → scrivi l’intera riga nel file di output.",
            AutoSize = true,
            Font = new Font("Segoe UI", 12.5f, FontStyle.Bold),
            ForeColor = Color.Yellow,
            Margin = new Padding(0, 0, 0, 12)
        };
        root.Controls.Add(lblTitle, 0, 0);
        root.SetColumnSpan(lblTitle, 6);

        // File A
        lblFileA = new Label { Text = "File A (.xlsx):", AutoSize = true, Margin = new Padding(0, 8, 8, 0) };
        root.Controls.Add(lblFileA, 0, 1);

        txtFileA = new TextBox
        {
            BackColor = Color.FromArgb(20, 20, 24),
            ForeColor = Color.Lime,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 6, 8, 0),
            Anchor = AnchorStyles.Left | AnchorStyles.Right
        };
        root.Controls.Add(txtFileA, 1, 1);

        btnBrowseA = new Button
        {
            Text = "Sfoglia…",
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 0)
        };
        root.Controls.Add(btnBrowseA, 2, 1);

        lblColsA = new Label
        {
            Text = "Scegli le DUE colonne dal File A (chiave composta):",
            AutoSize = true,
            Margin = new Padding(0, 10, 0, 4)
        };
        root.Controls.Add(lblColsA, 4, 1);
        root.SetColumnSpan(lblColsA, 2);

        lblA1 = new Label { Text = "Colonna A1:", AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        root.Controls.Add(lblA1, 4, 2);
        lblA2 = new Label { Text = "Colonna A2:", AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        root.Controls.Add(lblA2, 5, 2);

        lstACol1 = new ListBox
        {
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = Color.Lime,
            BorderStyle = BorderStyle.FixedSingle,
            Height = 180,
            SelectionMode = SelectionMode.One
        };
        root.Controls.Add(lstACol1, 4, 3);

        lstACol2 = new ListBox
        {
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = Color.Lime,
            BorderStyle = BorderStyle.FixedSingle,
            Height = 180,
            SelectionMode = SelectionMode.One
        };
        root.Controls.Add(lstACol2, 5, 3);

        // File B
        lblFileB = new Label { Text = "File B (.xlsx):", AutoSize = true, Margin = new Padding(0, 16, 8, 0) };
        root.Controls.Add(lblFileB, 0, 4);

        txtFileB = new TextBox
        {
            BackColor = Color.FromArgb(20, 20, 24),
            ForeColor = Color.Lime,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 14, 8, 0),
            Anchor = AnchorStyles.Left | AnchorStyles.Right
        };
        root.Controls.Add(txtFileB, 1, 4);

        btnBrowseB = new Button
        {
            Text = "Sfoglia…",
            AutoSize = true,
            Margin = new Padding(0, 12, 0, 0)
        };
        root.Controls.Add(btnBrowseB, 2, 4);

        lblColsB = new Label
        {
            Text = "Scegli le DUE colonne dal File B (chiave composta):",
            AutoSize = true,
            Margin = new Padding(0, 16, 0, 4)
        };
        root.Controls.Add(lblColsB, 4, 4);
        root.SetColumnSpan(lblColsB, 2);

        lblB1 = new Label { Text = "Colonna B1:", AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        root.Controls.Add(lblB1, 4, 5);
        lblB2 = new Label { Text = "Colonna B2:", AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        root.Controls.Add(lblB2, 5, 5);

        lstBCol1 = new ListBox
        {
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = Color.Lime,
            BorderStyle = BorderStyle.FixedSingle,
            Height = 180,
            SelectionMode = SelectionMode.One
        };
        root.Controls.Add(lstBCol1, 4, 6);

        lstBCol2 = new ListBox
        {
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = Color.Lime,
            BorderStyle = BorderStyle.FixedSingle,
            Height = 180,
            SelectionMode = SelectionMode.One
        };
        root.Controls.Add(lstBCol2, 5, 6);

        // Radio: da quale file estrarre le righe senza match
        grpOutputFrom = new GroupBox
        {
            Text = "Righe senza match da scrivere nell'output",
            AutoSize = true,
            ForeColor = Color.Lime,
            Margin = new Padding(0, 12, 0, 0)
        };
        rdoOutFromA = new RadioButton { Text = "Scrivi righe del File A che non trovano match in B", AutoSize = true, Checked = false, Margin = new Padding(8, 8, 8, 4) };
        rdoOutFromB = new RadioButton { Text = "Scrivi righe del File B che non trovano match in A", AutoSize = true, Checked = true, Margin = new Padding(8, 4, 8, 8) };
        grpOutputFrom.Controls.Add(rdoOutFromA);
        grpOutputFrom.Controls.Add(rdoOutFromB);
        root.Controls.Add(grpOutputFrom, 0, 7);
        root.SetColumnSpan(grpOutputFrom, 6);

        // Azioni
        pnlActions = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 10, 0, 0)
        };
        btnLoadHeaders = new Button { Text = "Carica intestazioni", AutoSize = true };
        btnRun = new Button { Text = "Avvia confronto e salva…", AutoSize = true, Margin = new Padding(12, 0, 0, 0) };
        pnlActions.Controls.AddRange(new Control[] { btnLoadHeaders, btnRun });
        root.Controls.Add(pnlActions, 0, 8);
        root.SetColumnSpan(pnlActions, 6);

        lblStatus = new Label
        {
            Text = "Pronto.",
            AutoSize = true,
            ForeColor = Color.Silver,
            Margin = new Padding(0, 6, 0, 0)
        };
        root.Controls.Add(lblStatus, 0, 9);
        root.SetColumnSpan(lblStatus, 6);
    }
}
