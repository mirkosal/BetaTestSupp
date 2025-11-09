using System.Drawing;
using System.Windows.Forms;

public partial class EditFieldForm : Form
{
    private System.ComponentModel.IContainer components = null;

    private TableLayoutPanel root;
    private Label lblTitle;          // Titolo in alto: "Rinomina campo — ..."
    private TableLayoutPanel row;    // Riga: [VecchioNome] [→] [TextBox NuovoNome]
    private Label lblOldName;
    private Label lblArrow;
    private TextBox txtValue;
    private FlowLayoutPanel pnlButtons;
    private Button btnOk;
    private Button btnCancel;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
            components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();

        // ===== Form =====
        SuspendLayout();
        Name = "EditFieldForm";
        Text = "Rinomina campo";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(24, 24, 28);
        ForeColor = Color.Lime;
        AutoScaleMode = AutoScaleMode.None;
        ClientSize = new Size(560, 168);
        Opacity = 0.96;

        // ===== root layout =====
        root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(24, 24, 28),
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(14, 12, 14, 12)
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));  // titolo
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // riga input
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));  // bottoni
        Controls.Add(root);

        // ===== Titolo =====
        lblTitle = new Label
        {
            Dock = DockStyle.Fill,
            ForeColor = Color.Khaki,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            Text = "Rinomina campo",
            AutoEllipsis = true
        };
        root.Controls.Add(lblTitle, 0, 0);

        // ===== riga vecchio → nuovo =====
        row = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3
        };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45)); // vecchio nome
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 28)); // freccia
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55)); // textbox nuovo
        root.Controls.Add(row, 0, 1);

        lblOldName = new Label
        {
            Dock = DockStyle.Fill,
            Text = "(vecchio)",
            TextAlign = ContentAlignment.MiddleRight,
            AutoEllipsis = true,
            Font = new Font("Segoe UI", 9F, FontStyle.Regular),
        };
        row.Controls.Add(lblOldName, 0, 0);

        lblArrow = new Label
        {
            Dock = DockStyle.Fill,
            Text = "→",
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 11F, FontStyle.Bold)
        };
        row.Controls.Add(lblArrow, 1, 0);

        txtValue = new TextBox
        {
            Dock = DockStyle.Fill,
            MaxLength = 128
        };
        row.Controls.Add(txtValue, 2, 0);

        // ===== Bottoni =====
        pnlButtons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 6, 0, 0)
        };
        root.Controls.Add(pnlButtons, 0, 2);

        btnCancel = new Button
        {
            Text = "Annulla",
            DialogResult = DialogResult.Cancel,
            Width = 96,
            Height = 28,
            Margin = new Padding(6, 6, 0, 0)
        };
        pnlButtons.Controls.Add(btnCancel);

        btnOk = new Button
        {
            Text = "Conferma",
            Width = 96,
            Height = 28,
            Margin = new Padding(6, 6, 6, 0)
        };
        pnlButtons.Controls.Add(btnOk);

        AcceptButton = btnOk;
        CancelButton = btnCancel;

        ResumeLayout(false);
    }
}
