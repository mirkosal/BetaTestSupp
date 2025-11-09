// File: CfCourseGeneratorForm.Designer.cs
using System.Drawing;
using System.Windows.Forms;

public partial class CfCourseGeneratorForm : BaseMenuForm
{
    private Panel pnlScroll;

    // ===== Fase 1 (come già consegnato in precedenza, invariata nella sostanza) =====
    private GroupBox grpPhase1;
    private Label lblTitle1, lblPhase1What, lblSheet, lblColumn, lblPhase1Hint, lblHint;
    private TextBox txtExcel, txtOutputFolder;
    private Button btnBrowse, btnOutBrowse, btnGenerate;
    private CheckBox chkSameFolder;
    private ComboBox cmbSheet, cmbColumn;
    private ListBox lstLog;
    private ProgressBar pgbPhase1;

    // ===== Fase 2 (come già consegnato, riassunto) =====
    private GroupBox grpPhase2;
    private Label lblTitle2, lblPhase2What, lblPhase2Note;
    private Label lblCourseMap, lblCourseSheet, lblCourseCol1, lblCourseCol2;
    private TextBox txtCourseMapPath;
    private Button btnBrowseCourseMap;
    private RadioButton rdoCourseSame, rdoCourseOther;
    private ComboBox cmbCourseSheet, cmbCourseCol1, cmbCourseCol2;

    private Label lblPersonMap, lblPersonSheet, lblPersonCol1, lblPersonCol2;
    private TextBox txtPersonMapPath;
    private Button btnBrowsePersonMap;
    private RadioButton rdoPersonSame, rdoPersonOther;
    private ComboBox cmbPersonSheet, cmbPersonCol1, cmbPersonCol2;

    // ===== Fase 3 (come già consegnato, con fix sovrapposizioni) =====
    private GroupBox grpPhase3;
    private Label lblTitle3, lblPhase3What, lblLiteralHint, lblTemplateName, lblPhase3Out, lblPreview3;
    private GroupBox grpTokens, grpFieldBuilder, grpTemplate;
    private ListBox lstTokens, lstTemplateFields, lstSavedTemplates;
    private TextBox txtLiteral, txtTemplateName, txtPhase3Out;
    private Button btnAddLiteral, btnClearField, btnAddFieldToTemplate, btnRemoveField,
                   btnSaveTemplate, btnLoadTemplate, btnDeleteTemplate,
                   btnGeneratePhase3, btnBrowsePhase3Out,
                   btnPreviewPhase3, btnOpenPhase3Folder,
                   btnRenameField;
    private FlowLayoutPanel pnlFieldChips;
    private ProgressBar pgbPhase3;
    private DataGridView dgvPhase3Preview;

    // ===== Fase 4 (nuova) =====
    private GroupBox grpPhase4;
    private Label lblTitle4, lblPhase4What, lblPhase4Out, lblPhase4Rules;
    private CheckedListBox chkRules;
    private Button btnAddRule, btnEditRule, btnDeleteRule, btnApplyPhase4, btnBrowsePhase4Out, btnOpenPhase4Folder, btnReloadFase3Headers;
    private TextBox txtPhase4Out;
    private ProgressBar pgbPhase4;

    private ToolTip tip;

    private void InitializeComponent()
    {
        var host = this.ContentHost;
        tip = new ToolTip();

        pnlScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        host.Controls.Add(pnlScroll);

        int x = 12, y = 12, w = 920;

        // ==== Fase 1 (omissis: stessa struttura fornita nel messaggio precedente) ====
        // Per brevità, qui si sottintende l’inizializzazione identica alla tua versione funzionante (campi, bottoni, progress ecc.)
        // ...
        // (Mantieni gli stessi handler già agganciati nel .cs)

        // ==== Fase 2 (omissis) ====
        // ...

        // ==== Fase 3 (omissis) ====
        // ...

        // ==== Fase 4 — Regole di esclusione ====
        grpPhase4 = new GroupBox { Text = "Fase 4 — Regole di esclusione", Location = new Point(x, y + 860), Size = new Size(w, 420) };
        lblTitle4 = new Label { AutoSize = true, Font = new Font("Segoe UI", 11F, FontStyle.Bold), Text = "Seleziona/gestisci regole da applicare all’output di Fase 3" };
        lblTitle4.Left = 12; lblTitle4.Top = 20;

        lblPhase4What = new Label
        {
            AutoSize = true,
            Left = 12,
            Top = 44,
            ForeColor = Color.DimGray,
            Text = "Spunta le regole attive. Puoi aggiungere/modificare una regola. L’applicazione genera un nuovo XLSX pulito e un CSV degli esclusi con il nome della regola come motivo."
        };

        // Output Fase 4
        lblPhase4Out = new Label { Left = 12, Top = 74, AutoSize = true, Text = "Output Fase 4 (.xlsx):" };
        txtPhase4Out = new TextBox { Left = 150, Top = 70, Width = 480, MaxLength = 260 };
        tip.SetToolTip(txtPhase4Out, "File di output Fase 4");

        btnBrowsePhase4Out = new Button { Left = txtPhase4Out.Right + 10, Top = 68, Width = 120, Text = "Sfoglia…" };
        btnOpenPhase4Folder = new Button { Left = btnBrowsePhase4Out.Right + 10, Top = 68, Width = 120, Text = "Apri cartella" };
        btnReloadFase3Headers = new Button { Left = btnOpenPhase4Folder.Right + 10, Top = 68, Width = 140, Text = "Rileggi intestazioni F3" };

        // Lista regole (checkbox)
        lblPhase4Rules = new Label { Left = 12, Top = 104, AutoSize = true, Text = "Regole disponibili:" };
        chkRules = new CheckedListBox
        {
            Left = 12,
            Top = 124,
            Width = 460,
            Height = 200,
            CheckOnClick = true
        };

        btnAddRule = new Button { Left = 484, Top = 124, Width = 120, Text = "Aggiungi…" };
        btnEditRule = new Button { Left = 484, Top = 158, Width = 120, Text = "Modifica…" };
        btnDeleteRule = new Button { Left = 484, Top = 192, Width = 120, Text = "Elimina" };

        pgbPhase4 = new ProgressBar
        {
            Left = 12,
            Top = 334,
            Width = grpPhase4.Width - 24,
            Height = 14,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Visible = false
        };
        btnApplyPhase4 = new Button { Left = grpPhase4.Width - 12 - 160, Top = 354, Width = 160, Text = "Applica Fase 4", Font = new Font("Segoe UI", 10f, FontStyle.Bold), Anchor = AnchorStyles.Top | AnchorStyles.Right };

        grpPhase4.Controls.AddRange(new Control[]
        {
            lblTitle4, lblPhase4What,
            lblPhase4Out, txtPhase4Out, btnBrowsePhase4Out, btnOpenPhase4Folder, btnReloadFase3Headers,
            lblPhase4Rules, chkRules, btnAddRule, btnEditRule, btnDeleteRule,
            pgbPhase4, btnApplyPhase4
        });
        pnlScroll.Controls.Add(grpPhase4);

        host.Resize += (_, __) =>
        {
            int inner = host.Width - 2 * x;
            grpPhase4.Width = inner;
            pgbPhase4.Width = grpPhase4.Width - 24;
            btnApplyPhase4.Left = grpPhase4.Width - 12 - btnApplyPhase4.Width;

            // Mantieni i bottoni allineati accanto alle textbox
            btnBrowsePhase4Out.Left = txtPhase4Out.Right + 10;
            btnOpenPhase4Folder.Left = btnBrowsePhase4Out.Right + 10;
            btnReloadFase3Headers.Left = btnOpenPhase4Folder.Right + 10;
        };
    }
}
