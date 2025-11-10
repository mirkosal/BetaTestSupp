// RuleEditorForm.Designer.cs — Designer minimale per l'editor regole (SENZA namespace)

#nullable enable
using System.Windows.Forms;
using System.Drawing;

public partial class RuleEditorForm
{
    private System.ComponentModel.IContainer components = null!;
    private SplitContainer split;
    private GroupBox grpList;
    private ListBox lstRules;
    private FlowLayoutPanel pnlListButtons;
    private Button btnNew;
    private Button btnDelete;

    private GroupBox grpEdit;
    private TableLayoutPanel tbl;
    private Label lblName;
    private Label lblType;
    private Label lblExpr;
    private TextBox txtName;
    private TextBox txtType;
    private TextBox txtExpr;
    private FlowLayoutPanel pnlEditButtons;
    private Button btnSave;
    private Button btnClose;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null)) components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();

        split = new SplitContainer();
        grpList = new GroupBox();
        lstRules = new ListBox();
        pnlListButtons = new FlowLayoutPanel();
        btnNew = new Button();
        btnDelete = new Button();

        grpEdit = new GroupBox();
        tbl = new TableLayoutPanel();
        lblName = new Label();
        lblType = new Label();
        lblExpr = new Label();
        txtName = new TextBox();
        txtType = new TextBox();
        txtExpr = new TextBox();
        pnlEditButtons = new FlowLayoutPanel();
        btnSave = new Button();
        btnClose = new Button();

        // Form
        this.SuspendLayout();
        this.Text = "Editor Regole (Fase 4)";
        this.StartPosition = FormStartPosition.CenterParent;
        this.ClientSize = new Size(820, 520);
        this.MinimumSize = new Size(760, 440);

        // Split
        split.Dock = DockStyle.Fill;
        split.Orientation = Orientation.Vertical;
        split.SplitterDistance = 280;
        split.Panel1MinSize = 240;
        split.Panel2MinSize = 400;
        this.Controls.Add(split);

        // Left: List
        grpList.Text = "Regole";
        grpList.Dock = DockStyle.Fill;

        lstRules.Dock = DockStyle.Fill;
        lstRules.IntegralHeight = false;

        pnlListButtons.Dock = DockStyle.Bottom;
        pnlListButtons.Height = 42;
        pnlListButtons.FlowDirection = FlowDirection.RightToLeft;
        pnlListButtons.Padding = new Padding(6);

        btnNew.Text = "Nuova";
        btnNew.Width = 100;
        btnNew.Height = 28;
        btnDelete.Text = "Elimina";
        btnDelete.Width = 100;
        btnDelete.Height = 28;

        pnlListButtons.Controls.Add(btnNew);
        pnlListButtons.Controls.Add(btnDelete);

        grpList.Controls.Add(lstRules);
        grpList.Controls.Add(pnlListButtons);

        split.Panel1.Controls.Add(grpList);

        // Right: Edit
        grpEdit.Text = "Dettaglio";
        grpEdit.Dock = DockStyle.Fill;

        tbl.ColumnCount = 2;
        tbl.RowCount = 4;
        tbl.Dock = DockStyle.Fill;
        tbl.Padding = new Padding(8);
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));

        lblName.Text = "Nome:";
        lblName.TextAlign = ContentAlignment.MiddleRight;
        lblName.Dock = DockStyle.Fill;

        lblType.Text = "Tipo:";
        lblType.TextAlign = ContentAlignment.MiddleRight;
        lblType.Dock = DockStyle.Fill;

        lblExpr.Text = "Espressione:";
        lblExpr.TextAlign = ContentAlignment.MiddleRight;
        lblExpr.Dock = DockStyle.Fill;

        txtName.Dock = DockStyle.Fill;
        txtType.Dock = DockStyle.Fill;
        txtExpr.Dock = DockStyle.Fill;
        txtExpr.Multiline = true;
        txtExpr.ScrollBars = ScrollBars.Vertical;

        pnlEditButtons.Dock = DockStyle.Fill;
        pnlEditButtons.FlowDirection = FlowDirection.RightToLeft;

        btnSave.Text = "Salva";
        btnSave.Width = 110;
        btnSave.Height = 30;

        btnClose.Text = "Chiudi";
        btnClose.Width = 110;
        btnClose.Height = 30;

        pnlEditButtons.Controls.Add(btnSave);
        pnlEditButtons.Controls.Add(btnClose);

        tbl.Controls.Add(lblName, 0, 0);
        tbl.Controls.Add(txtName, 1, 0);
        tbl.Controls.Add(lblType, 0, 1);
        tbl.Controls.Add(txtType, 1, 1);
        tbl.Controls.Add(lblExpr, 0, 2);
        tbl.Controls.Add(txtExpr, 1, 2);
        tbl.Controls.Add(pnlEditButtons, 0, 3);
        tbl.SetColumnSpan(pnlEditButtons, 2);

        grpEdit.Controls.Add(tbl);
        split.Panel2.Controls.Add(grpEdit);

        this.ResumeLayout(false);
    }
}
