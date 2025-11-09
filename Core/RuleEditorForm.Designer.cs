using System.Drawing;
using System.Windows.Forms;
using BetaTestSupp.Core;

public partial class RuleEditorForm : Form
{
    private System.ComponentModel.IContainer components = null;

    private TableLayoutPanel root;
    private Label lblTitle;

    private Label lblName;
    private TextBox txtName;

    private Label lblKind;
    private ComboBox cmbKind;

    private Label lblDataset;
    private ComboBox cmbDataset;

    private Label lblField1;
    private ComboBox cmbField1;

    private Label lblField2;
    private ComboBox cmbField2;

    private Label lblIntParam;
    private NumericUpDown numIntParam;

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
        SuspendLayout();

        Text = "Regola";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false; MaximizeBox = false; ShowInTaskbar = false;
        ClientSize = new Size(560, 300);

        root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 8,
            Padding = new Padding(12)
        };
        for (int i = 0; i < 4; i++) root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28)); // titolo
        for (int i = 0; i < 6; i++) root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // spacer
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // bottoni
        Controls.Add(root);

        lblTitle = new Label { Text = "Nuova regola", AutoSize = true, Font = new Font("Segoe UI", 10.5f, FontStyle.Bold) };
        root.Controls.Add(lblTitle, 0, 0);
        root.SetColumnSpan(lblTitle, 4);

        lblName = new Label { Text = "Nome regola:", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
        txtName = new TextBox { Dock = DockStyle.Fill, MaxLength = 128 };
        root.Controls.Add(lblName, 0, 1); root.SetColumnSpan(lblName, 1);
        root.Controls.Add(txtName, 1, 1); root.SetColumnSpan(txtName, 3);

        lblKind = new Label { Text = "Tipo:", AutoSize = true };
        cmbKind = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        root.Controls.Add(lblKind, 0, 2);
        root.Controls.Add(cmbKind, 1, 2); root.SetColumnSpan(cmbKind, 3);

        lblDataset = new Label { Text = "Dataset Fase 2:", AutoSize = true };
        cmbDataset = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        cmbDataset.Items.AddRange(new object[] { "", "Persone", "Corsi" });
        root.Controls.Add(lblDataset, 0, 3);
        root.Controls.Add(cmbDataset, 1, 3); root.SetColumnSpan(cmbDataset, 3);

        lblField1 = new Label { Text = "Campo 1:", AutoSize = true };
        cmbField1 = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        root.Controls.Add(lblField1, 0, 4);
        root.Controls.Add(cmbField1, 1, 4); root.SetColumnSpan(cmbField1, 3);

        lblField2 = new Label { Text = "Campo 2 (opz.):", AutoSize = true };
        cmbField2 = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        root.Controls.Add(lblField2, 0, 5);
        root.Controls.Add(cmbField2, 1, 5); root.SetColumnSpan(cmbField2, 3);

        lblIntParam = new Label { Text = "Parametro intero:", AutoSize = true };
        numIntParam = new NumericUpDown { Minimum = 0, Maximum = 1000, Value = 16, Dock = DockStyle.Left, Width = 100 };
        root.Controls.Add(lblIntParam, 0, 6);
        root.Controls.Add(numIntParam, 1, 6);

        pnlButtons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
        btnCancel = new Button { Text = "Annulla", DialogResult = DialogResult.Cancel, Width = 90 };
        btnOk = new Button { Text = "Conferma", Width = 90 };
        pnlButtons.Controls.Add(btnCancel);
        pnlButtons.Controls.Add(btnOk);
        root.Controls.Add(pnlButtons, 0, 7);
        root.SetColumnSpan(pnlButtons, 4);

        AcceptButton = btnOk;
        CancelButton = btnCancel;

        ResumeLayout(false);
    }
}
