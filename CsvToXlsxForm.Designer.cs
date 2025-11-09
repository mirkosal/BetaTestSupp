using System.Drawing;
using System.Windows.Forms;

public partial class CsvToXlsxForm : BaseMenuForm
{
    private System.ComponentModel.IContainer components = null;

    private TableLayoutPanel root;
    private Label lblTitle;
    private Label lblCsv;
    private TextBox txtCsvPath;
    private Button btnBrowse;

    private Label lblSep;
    private TextBox txtSep;
    private Label lblEsc;
    private TextBox txtEsc;

    private CheckBox chkHeader;
    private CheckBox chkDetectAllDates;

    private Button btnConvert;
    private FlowLayoutPanel pnlButtons;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null)) components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();

        root = new TableLayoutPanel();
        lblTitle = new Label();
        lblCsv = new Label();
        txtCsvPath = new TextBox();
        btnBrowse = new Button();

        lblSep = new Label();
        txtSep = new TextBox();
        lblEsc = new Label();
        txtEsc = new TextBox();

        chkHeader = new CheckBox();
        chkDetectAllDates = new CheckBox();

        pnlButtons = new FlowLayoutPanel();
        btnConvert = new Button();

        SuspendLayout();
        root.SuspendLayout();

        // Form
        Text = "Converti CSV → XLSX";
        Width = 720;
        Height = 300;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(24, 24, 28);
        ForeColor = Color.Lime;
        FormBorderStyle = FormBorderStyle.FixedDialog;

        // Root layout (verrà aggiunto al ContentHost del BaseMenuForm)
        root.Dock = DockStyle.Fill;
        root.Padding = new Padding(12);
        root.ColumnCount = 6;
        root.RowCount = 5;
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));

        // === Aggiunta al ContentHost del base ===
        // (Così non c'è sovrapposizione col menu e tutto è raggruppato)
        ContentHost.Controls.Add(root);

        // Titolo
        lblTitle.Text = "Conversione CSV → XLSX";
        lblTitle.AutoSize = true;
        lblTitle.Font = new Font("Segoe UI", 13F, FontStyle.Bold);
        lblTitle.ForeColor = Color.Yellow;
        root.Controls.Add(lblTitle, 0, 0);
        root.SetColumnSpan(lblTitle, 6);

        // CSV path
        lblCsv.Text = "File CSV:";
        lblCsv.AutoSize = true;
        lblCsv.Margin = new Padding(0, 14, 8, 0);
        root.Controls.Add(lblCsv, 0, 1);

        txtCsvPath.Margin = new Padding(0, 10, 8, 0);
        txtCsvPath.BackColor = Color.FromArgb(20, 20, 24);
        txtCsvPath.ForeColor = Color.Lime;
        txtCsvPath.BorderStyle = BorderStyle.FixedSingle;
        root.Controls.Add(txtCsvPath, 1, 1);
        root.SetColumnSpan(txtCsvPath, 3);

        btnBrowse.Text = "Sfoglia…";
        btnBrowse.Margin = new Padding(0, 8, 0, 0);
        btnBrowse.Click += btnBrowse_Click;
        root.Controls.Add(btnBrowse, 4, 1);
        root.SetColumnSpan(btnBrowse, 2);

        // Separatore / Escape
        lblSep.Text = "Separatore:";
        lblSep.AutoSize = true;
        lblSep.Margin = new Padding(0, 14, 8, 0);
        root.Controls.Add(lblSep, 0, 2);

        txtSep.Text = ";";
        txtSep.MaxLength = 1;
        txtSep.Width = 40;
        txtSep.Margin = new Padding(0, 10, 8, 0);
        txtSep.BackColor = Color.FromArgb(20, 20, 24);
        txtSep.ForeColor = Color.Lime;
        txtSep.TextChanged += EnforceOneChar;
        root.Controls.Add(txtSep, 1, 2);

        lblEsc.Text = "Escape:";
        lblEsc.AutoSize = true;
        lblEsc.Margin = new Padding(0, 14, 8, 0);
        root.Controls.Add(lblEsc, 2, 2);

        txtEsc.Text = "\"";
        txtEsc.MaxLength = 1;
        txtEsc.Width = 40;
        txtEsc.Margin = new Padding(0, 10, 8, 0);
        txtEsc.BackColor = Color.FromArgb(20, 20, 24);
        txtEsc.ForeColor = Color.Lime;
        txtEsc.TextChanged += EnforceOneChar;
        root.Controls.Add(txtEsc, 3, 2);

        chkHeader.Text = "Intestazione presente (riga 1)";
        chkHeader.AutoSize = true;
        chkHeader.Checked = true;
        chkHeader.Margin = new Padding(0, 10, 0, 0);
        root.Controls.Add(chkHeader, 4, 2);
        root.SetColumnSpan(chkHeader, 2);

        // Opzione: riconoscimento date su tutto il documento
        chkDetectAllDates.Text = "Riconosci automaticamente tutte le date (formato: yyyy-mm-dd)";
        chkDetectAllDates.AutoSize = true;
        chkDetectAllDates.Checked = true;
        chkDetectAllDates.Margin = new Padding(0, 12, 0, 0);
        root.Controls.Add(chkDetectAllDates, 0, 3);
        root.SetColumnSpan(chkDetectAllDates, 6);

        // Pulsanti
        pnlButtons.FlowDirection = FlowDirection.RightToLeft;
        pnlButtons.Dock = DockStyle.Fill;
        pnlButtons.Margin = new Padding(0, 12, 0, 0);

        btnConvert.Text = "Converti in XLSX…";
        btnConvert.AutoSize = true;
        btnConvert.Margin = new Padding(8, 8, 0, 0);
        btnConvert.Click += btnConvert_Click;

        pnlButtons.Controls.Add(btnConvert);
        root.Controls.Add(pnlButtons, 0, 4);
        root.SetColumnSpan(pnlButtons, 6);

        root.ResumeLayout(false);
        root.PerformLayout();
        ResumeLayout(false);
    }
}
