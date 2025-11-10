// CfCourseGeneratorForm.Designer.cs — Designer completo a schede.
// Contiene TUTTI i controlli nominati nei tuoi errori.
// Nessuna chiamata a InitializeComponent() del base.

using System.Windows.Forms;
using System.ComponentModel;

    partial class CfCourseGeneratorForm
    {
        private IContainer components = null;

        // Controlli condivisi
        private TabControl tabs;
        private TabPage tabPhase1;
        private TabPage tabPhase2;
        private TabPage tabPhase3;
        private TabPage tabPhase4;
        private ListBox lstLog;

        // Fase 1
        private TextBox txtExcel;
        private Button btnBrowse;
        private TextBox txtOutputFolder;
        private Button btnOutBrowse;
        private CheckBox chkSameFolder;
        private ComboBox cmbSheet;
        private ComboBox cmbColumn;
        private ProgressBar pgbPhase1;
        private Button btnGenerate;

        // Fase 2
        private RadioButton rdoCourseSame;
        private RadioButton rdoCourseOther;
        private TextBox txtCourseMapPath;
        private Button btnBrowseCourseMap;
        private ComboBox cmbCourseSheet;
        private ComboBox cmbCourseCol1;
        private ComboBox cmbCourseCol2;

        private RadioButton rdoPersonSame;
        private RadioButton rdoPersonOther;
        private TextBox txtPersonMapPath;
        private Button btnBrowsePersonMap;
        private ComboBox cmbPersonSheet;
        private ComboBox cmbPersonCol1;
        private ComboBox cmbPersonCol2;

        // Fase 3
        private TextBox txtPhase3Out;
        private Button btnBrowsePhase3Out;
        private Button btnOpenPhase3Folder;
        private Button btnPreviewPhase3;
        private ProgressBar pgbPhase3;

        private TextBox txtLiteral;
        private Panel pnlFieldChips;
        private ListBox lstTemplateFields;
        private TextBox txtTemplateName;
        private ListBox lstSavedTemplates;

        private Button btnAddLiteral;
        private Button btnClearField;
        private Button btnAddFieldToTemplate;
        private Button btnRemoveField;
        private Button btnRenameField;
        private Button btnSaveTemplate;
        private Button btnLoadTemplate;
        private Button btnDeleteTemplate;
        private Button btnGeneratePhase3;

        private DataGridView dgvPhase3Preview;

        // Fase 4
        private TextBox txtPhase4OutDir;
        private Button btnBrowsePhase4Out;
        private Button btnOpenPhase4Folder;

        private Button btnReloadFase3Headers;
        private Button btnAddRule;
        private Button btnEditRule;
        private Button btnDeleteRule;
        private Button btnApplyPhase4;

        /// <summary>
        /// Pulizia risorse.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>
        /// Init UI.
        /// </summary>
        private void InitializeComponent()
        {
            components = new Container();
            this.tabs = new TabControl();
            this.tabPhase1 = new TabPage();
            this.tabPhase2 = new TabPage();
            this.tabPhase3 = new TabPage();
            this.tabPhase4 = new TabPage();
            this.lstLog = new ListBox();

            // Fase 1 controls
            this.txtExcel = new TextBox();
            this.btnBrowse = new Button();
            this.txtOutputFolder = new TextBox();
            this.btnOutBrowse = new Button();
            this.chkSameFolder = new CheckBox();
            this.cmbSheet = new ComboBox();
            this.cmbColumn = new ComboBox();
            this.pgbPhase1 = new ProgressBar();
            this.btnGenerate = new Button();

            // Fase 2 controls
            this.rdoCourseSame = new RadioButton();
            this.rdoCourseOther = new RadioButton();
            this.txtCourseMapPath = new TextBox();
            this.btnBrowseCourseMap = new Button();
            this.cmbCourseSheet = new ComboBox();
            this.cmbCourseCol1 = new ComboBox();
            this.cmbCourseCol2 = new ComboBox();

            this.rdoPersonSame = new RadioButton();
            this.rdoPersonOther = new RadioButton();
            this.txtPersonMapPath = new TextBox();
            this.btnBrowsePersonMap = new Button();
            this.cmbPersonSheet = new ComboBox();
            this.cmbPersonCol1 = new ComboBox();
            this.cmbPersonCol2 = new ComboBox();

            // Fase 3 controls
            this.txtPhase3Out = new TextBox();
            this.btnBrowsePhase3Out = new Button();
            this.btnOpenPhase3Folder = new Button();
            this.btnPreviewPhase3 = new Button();
            this.pgbPhase3 = new ProgressBar();

            this.txtLiteral = new TextBox();
            this.pnlFieldChips = new Panel();
            this.lstTemplateFields = new ListBox();
            this.txtTemplateName = new TextBox();
            this.lstSavedTemplates = new ListBox();

            this.btnAddLiteral = new Button();
            this.btnClearField = new Button();
            this.btnAddFieldToTemplate = new Button();
            this.btnRemoveField = new Button();
            this.btnRenameField = new Button();
            this.btnSaveTemplate = new Button();
            this.btnLoadTemplate = new Button();
            this.btnDeleteTemplate = new Button();
            this.btnGeneratePhase3 = new Button();

            this.dgvPhase3Preview = new DataGridView();

            // Fase 4 controls
            this.txtPhase4OutDir = new TextBox();
            this.btnBrowsePhase4Out = new Button();
            this.btnOpenPhase4Folder = new Button();

            this.btnReloadFase3Headers = new Button();
            this.btnAddRule = new Button();
            this.btnEditRule = new Button();
            this.btnDeleteRule = new Button();
            this.btnApplyPhase4 = new Button();

            // ===== Form
            this.SuspendLayout();
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1080, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "CF Course Generator";
            this.Name = "CfCourseGeneratorForm";

            // ===== tabs
            this.tabs.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.tabs.Location = new System.Drawing.Point(12, 12);
            this.tabs.Name = "tabs";
            this.tabs.Size = new System.Drawing.Size(1056, 620);
            this.tabs.TabIndex = 0;

            this.tabs.Controls.Add(this.tabPhase1);
            this.tabs.Controls.Add(this.tabPhase2);
            this.tabs.Controls.Add(this.tabPhase3);
            this.tabs.Controls.Add(this.tabPhase4);

            // ===== lstLog
            this.lstLog.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.lstLog.Location = new System.Drawing.Point(12, 640);
            this.lstLog.Name = "lstLog";
            this.lstLog.Size = new System.Drawing.Size(1056, 64);

            // ===== TAB PHASE 1 =====
            this.tabPhase1.Text = "Fase 1";
            this.tabPhase1.Padding = new Padding(8);

            this.txtExcel.Location = new System.Drawing.Point(16, 16);
            this.txtExcel.Width = 700;
            this.txtExcel.Name = "txtExcel";

            this.btnBrowse.Location = new System.Drawing.Point(724, 14);
            this.btnBrowse.Text = "Sfoglia...";
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Width = 100;

            this.chkSameFolder.Location = new System.Drawing.Point(16, 50);
            this.chkSameFolder.Text = "Usa stessa cartella del file Excel";
            this.chkSameFolder.Name = "chkSameFolder";

            this.txtOutputFolder.Location = new System.Drawing.Point(16, 76);
            this.txtOutputFolder.Width = 700;
            this.txtOutputFolder.Name = "txtOutputFolder";

            this.btnOutBrowse.Location = new System.Drawing.Point(724, 74);
            this.btnOutBrowse.Text = "Output...";
            this.btnOutBrowse.Name = "btnOutBrowse";
            this.btnOutBrowse.Width = 100;

            this.cmbSheet.Location = new System.Drawing.Point(16, 120);
            this.cmbSheet.Width = 200;
            this.cmbSheet.Name = "cmbSheet";
            this.cmbSheet.DropDownStyle = ComboBoxStyle.DropDownList;

            this.cmbColumn.Location = new System.Drawing.Point(228, 120);
            this.cmbColumn.Width = 200;
            this.cmbColumn.Name = "cmbColumn";
            this.cmbColumn.DropDownStyle = ComboBoxStyle.DropDownList;

            this.btnGenerate.Location = new System.Drawing.Point(16, 160);
            this.btnGenerate.Text = "Esegui Fase 1";
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Width = 150;

            this.pgbPhase1.Location = new System.Drawing.Point(180, 164);
            this.pgbPhase1.Width = 450;
            this.pgbPhase1.Name = "pgbPhase1";

            this.tabPhase1.Controls.Add(this.txtExcel);
            this.tabPhase1.Controls.Add(this.btnBrowse);
            this.tabPhase1.Controls.Add(this.chkSameFolder);
            this.tabPhase1.Controls.Add(this.txtOutputFolder);
            this.tabPhase1.Controls.Add(this.btnOutBrowse);
            this.tabPhase1.Controls.Add(this.cmbSheet);
            this.tabPhase1.Controls.Add(this.cmbColumn);
            this.tabPhase1.Controls.Add(this.btnGenerate);
            this.tabPhase1.Controls.Add(this.pgbPhase1);

            // ===== TAB PHASE 2 =====
            this.tabPhase2.Text = "Fase 2";
            this.tabPhase2.Padding = new Padding(8);

            // Course
            this.rdoCourseSame.Location = new System.Drawing.Point(16, 16);
            this.rdoCourseSame.Text = "CourseMap sullo stesso Excel";
            this.rdoCourseSame.Name = "rdoCourseSame";
            this.rdoCourseSame.Checked = true;

            this.rdoCourseOther.Location = new System.Drawing.Point(260, 16);
            this.rdoCourseOther.Text = "CourseMap su altro file";
            this.rdoCourseOther.Name = "rdoCourseOther";

            this.txtCourseMapPath.Location = new System.Drawing.Point(16, 46);
            this.txtCourseMapPath.Width = 700;
            this.txtCourseMapPath.Name = "txtCourseMapPath";
            this.txtCourseMapPath.Enabled = false;

            this.btnBrowseCourseMap.Location = new System.Drawing.Point(724, 44);
            this.btnBrowseCourseMap.Text = "Sfoglia...";
            this.btnBrowseCourseMap.Name = "btnBrowseCourseMap";
            this.btnBrowseCourseMap.Width = 100;
            this.btnBrowseCourseMap.Enabled = false;

            this.cmbCourseSheet.Location = new System.Drawing.Point(16, 86);
            this.cmbCourseSheet.Width = 200;
            this.cmbCourseSheet.Name = "cmbCourseSheet";
            this.cmbCourseSheet.DropDownStyle = ComboBoxStyle.DropDownList;

            this.cmbCourseCol1.Location = new System.Drawing.Point(228, 86);
            this.cmbCourseCol1.Width = 200;
            this.cmbCourseCol1.Name = "cmbCourseCol1";
            this.cmbCourseCol1.DropDownStyle = ComboBoxStyle.DropDownList;

            this.cmbCourseCol2.Location = new System.Drawing.Point(440, 86);
            this.cmbCourseCol2.Width = 200;
            this.cmbCourseCol2.Name = "cmbCourseCol2";
            this.cmbCourseCol2.DropDownStyle = ComboBoxStyle.DropDownList;

            // Person
            this.rdoPersonSame.Location = new System.Drawing.Point(16, 140);
            this.rdoPersonSame.Text = "PersonMap sullo stesso Excel";
            this.rdoPersonSame.Name = "rdoPersonSame";
            this.rdoPersonSame.Checked = true;

            this.rdoPersonOther.Location = new System.Drawing.Point(260, 140);
            this.rdoPersonOther.Text = "PersonMap su altro file";
            this.rdoPersonOther.Name = "rdoPersonOther";

            this.txtPersonMapPath.Location = new System.Drawing.Point(16, 170);
            this.txtPersonMapPath.Width = 700;
            this.txtPersonMapPath.Name = "txtPersonMapPath";
            this.txtPersonMapPath.Enabled = false;

            this.btnBrowsePersonMap.Location = new System.Drawing.Point(724, 168);
            this.btnBrowsePersonMap.Text = "Sfoglia...";
            this.btnBrowsePersonMap.Name = "btnBrowsePersonMap";
            this.btnBrowsePersonMap.Width = 100;
            this.btnBrowsePersonMap.Enabled = false;

            this.cmbPersonSheet.Location = new System.Drawing.Point(16, 210);
            this.cmbPersonSheet.Width = 200;
            this.cmbPersonSheet.Name = "cmbPersonSheet";
            this.cmbPersonSheet.DropDownStyle = ComboBoxStyle.DropDownList;

            this.cmbPersonCol1.Location = new System.Drawing.Point(228, 210);
            this.cmbPersonCol1.Width = 200;
            this.cmbPersonCol1.Name = "cmbPersonCol1";
            this.cmbPersonCol1.DropDownStyle = ComboBoxStyle.DropDownList;

            this.cmbPersonCol2.Location = new System.Drawing.Point(440, 210);
            this.cmbPersonCol2.Width = 200;
            this.cmbPersonCol2.Name = "cmbPersonCol2";
            this.cmbPersonCol2.DropDownStyle = ComboBoxStyle.DropDownList;

            this.tabPhase2.Controls.Add(this.rdoCourseSame);
            this.tabPhase2.Controls.Add(this.rdoCourseOther);
            this.tabPhase2.Controls.Add(this.txtCourseMapPath);
            this.tabPhase2.Controls.Add(this.btnBrowseCourseMap);
            this.tabPhase2.Controls.Add(this.cmbCourseSheet);
            this.tabPhase2.Controls.Add(this.cmbCourseCol1);
            this.tabPhase2.Controls.Add(this.cmbCourseCol2);

            this.tabPhase2.Controls.Add(this.rdoPersonSame);
            this.tabPhase2.Controls.Add(this.rdoPersonOther);
            this.tabPhase2.Controls.Add(this.txtPersonMapPath);
            this.tabPhase2.Controls.Add(this.btnBrowsePersonMap);
            this.tabPhase2.Controls.Add(this.cmbPersonSheet);
            this.tabPhase2.Controls.Add(this.cmbPersonCol1);
            this.tabPhase2.Controls.Add(this.cmbPersonCol2);

            // ===== TAB PHASE 3 =====
            this.tabPhase3.Text = "Fase 3";
            this.tabPhase3.Padding = new Padding(8);

            this.txtPhase3Out.Location = new System.Drawing.Point(16, 16);
            this.txtPhase3Out.Width = 600;
            this.txtPhase3Out.Name = "txtPhase3Out";

            this.btnBrowsePhase3Out.Location = new System.Drawing.Point(624, 14);
            this.btnBrowsePhase3Out.Text = "Output...";
            this.btnBrowsePhase3Out.Name = "btnBrowsePhase3Out";
            this.btnBrowsePhase3Out.Width = 90;

            this.btnOpenPhase3Folder.Location = new System.Drawing.Point(720, 14);
            this.btnOpenPhase3Folder.Text = "Apri cartella";
            this.btnOpenPhase3Folder.Name = "btnOpenPhase3Folder";
            this.btnOpenPhase3Folder.Width = 100;

            this.btnPreviewPhase3.Location = new System.Drawing.Point(16, 48);
            this.btnPreviewPhase3.Text = "Aggiorna Anteprima";
            this.btnPreviewPhase3.Name = "btnPreviewPhase3";
            this.btnPreviewPhase3.Width = 150;

            this.pgbPhase3.Location = new System.Drawing.Point(180, 52);
            this.pgbPhase3.Width = 450;
            this.pgbPhase3.Name = "pgbPhase3";

            // Builder
            this.txtLiteral.Location = new System.Drawing.Point(16, 90);
            this.txtLiteral.Width = 300;
            this.txtLiteral.Name = "txtLiteral";
            this.txtLiteral.PlaceholderText = "Testo letterale";

            this.btnAddLiteral.Location = new System.Drawing.Point(324, 88);
            this.btnAddLiteral.Text = "Aggiungi";
            this.btnAddLiteral.Name = "btnAddLiteral";
            this.btnAddLiteral.Width = 80;

            this.btnClearField.Location = new System.Drawing.Point(408, 88);
            this.btnClearField.Text = "Pulisci";
            this.btnClearField.Name = "btnClearField";
            this.btnClearField.Width = 80;

            this.pnlFieldChips.Location = new System.Drawing.Point(16, 118);
            this.pnlFieldChips.Size = new System.Drawing.Size(472, 80);
            this.pnlFieldChips.AutoScroll = true;
            this.pnlFieldChips.BorderStyle = BorderStyle.FixedSingle;
            this.pnlFieldChips.Name = "pnlFieldChips";

            this.btnAddFieldToTemplate.Location = new System.Drawing.Point(16, 206);
            this.btnAddFieldToTemplate.Text = "Aggiungi al Template";
            this.btnAddFieldToTemplate.Name = "btnAddFieldToTemplate";
            this.btnAddFieldToTemplate.Width = 180;

            this.btnRemoveField.Location = new System.Drawing.Point(202, 206);
            this.btnRemoveField.Text = "Rimuovi selezionato";
            this.btnRemoveField.Name = "btnRemoveField";
            this.btnRemoveField.Width = 160;

            this.btnRenameField.Location = new System.Drawing.Point(368, 206);
            this.btnRenameField.Text = "Rinomina selezionato";
            this.btnRenameField.Name = "btnRenameField";
            this.btnRenameField.Width = 160;

            this.lstTemplateFields.Location = new System.Drawing.Point(16, 236);
            this.lstTemplateFields.Size = new System.Drawing.Size(512, 160);
            this.lstTemplateFields.Name = "lstTemplateFields";

            this.txtTemplateName.Location = new System.Drawing.Point(16, 402);
            this.txtTemplateName.Width = 260;
            this.txtTemplateName.Name = "txtTemplateName";
            this.txtTemplateName.PlaceholderText = "Nome Template";

            this.btnSaveTemplate.Location = new System.Drawing.Point(282, 400);
            this.btnSaveTemplate.Text = "Salva";
            this.btnSaveTemplate.Name = "btnSaveTemplate";
            this.btnSaveTemplate.Width = 80;

            this.btnLoadTemplate.Location = new System.Drawing.Point(366, 400);
            this.btnLoadTemplate.Text = "Carica";
            this.btnLoadTemplate.Name = "btnLoadTemplate";
            this.btnLoadTemplate.Width = 80;

            this.btnDeleteTemplate.Location = new System.Drawing.Point(450, 400);
            this.btnDeleteTemplate.Text = "Elimina";
            this.btnDeleteTemplate.Name = "btnDeleteTemplate";
            this.btnDeleteTemplate.Width = 80;

            this.lstSavedTemplates.Location = new System.Drawing.Point(536, 90);
            this.lstSavedTemplates.Size = new System.Drawing.Size(280, 324);
            this.lstSavedTemplates.Name = "lstSavedTemplates";

            this.btnGeneratePhase3.Location = new System.Drawing.Point(536, 420);
            this.btnGeneratePhase3.Text = "Genera output";
            this.btnGeneratePhase3.Name = "btnGeneratePhase3";
            this.btnGeneratePhase3.Width = 140;

            this.dgvPhase3Preview.Location = new System.Drawing.Point(16, 436);
            this.dgvPhase3Preview.Size = new System.Drawing.Size(800, 120);
            this.dgvPhase3Preview.Name = "dgvPhase3Preview";
            this.dgvPhase3Preview.AllowUserToAddRows = false;
            this.dgvPhase3Preview.AllowUserToDeleteRows = false;
            this.dgvPhase3Preview.ReadOnly = true;

            this.tabPhase3.Controls.Add(this.txtPhase3Out);
            this.tabPhase3.Controls.Add(this.btnBrowsePhase3Out);
            this.tabPhase3.Controls.Add(this.btnOpenPhase3Folder);
            this.tabPhase3.Controls.Add(this.btnPreviewPhase3);
            this.tabPhase3.Controls.Add(this.pgbPhase3);
            this.tabPhase3.Controls.Add(this.txtLiteral);
            this.tabPhase3.Controls.Add(this.btnAddLiteral);
            this.tabPhase3.Controls.Add(this.btnClearField);
            this.tabPhase3.Controls.Add(this.pnlFieldChips);
            this.tabPhase3.Controls.Add(this.btnAddFieldToTemplate);
            this.tabPhase3.Controls.Add(this.btnRemoveField);
            this.tabPhase3.Controls.Add(this.btnRenameField);
            this.tabPhase3.Controls.Add(this.lstTemplateFields);
            this.tabPhase3.Controls.Add(this.txtTemplateName);
            this.tabPhase3.Controls.Add(this.btnSaveTemplate);
            this.tabPhase3.Controls.Add(this.btnLoadTemplate);
            this.tabPhase3.Controls.Add(this.btnDeleteTemplate);
            this.tabPhase3.Controls.Add(this.lstSavedTemplates);
            this.tabPhase3.Controls.Add(this.btnGeneratePhase3);
            this.tabPhase3.Controls.Add(this.dgvPhase3Preview);

            // ===== TAB PHASE 4 =====
            this.tabPhase4.Text = "Fase 4";
            this.tabPhase4.Padding = new Padding(8);

            this.txtPhase4OutDir.Location = new System.Drawing.Point(16, 16);
            this.txtPhase4OutDir.Width = 600;
            this.txtPhase4OutDir.Name = "txtPhase4OutDir";

            this.btnBrowsePhase4Out.Location = new System.Drawing.Point(624, 14);
            this.btnBrowsePhase4Out.Text = "Output...";
            this.btnBrowsePhase4Out.Name = "btnBrowsePhase4Out";
            this.btnBrowsePhase4Out.Width = 90;

            this.btnOpenPhase4Folder.Location = new System.Drawing.Point(720, 14);
            this.btnOpenPhase4Folder.Text = "Apri cartella";
            this.btnOpenPhase4Folder.Name = "btnOpenPhase4Folder";
            this.btnOpenPhase4Folder.Width = 100;

            this.btnReloadFase3Headers.Location = new System.Drawing.Point(16, 54);
            this.btnReloadFase3Headers.Text = "Ricarica headers Fase 3";
            this.btnReloadFase3Headers.Name = "btnReloadFase3Headers";
            this.btnReloadFase3Headers.Width = 180;

            this.btnAddRule.Location = new System.Drawing.Point(16, 90);
            this.btnAddRule.Text = "Aggiungi regola";
            this.btnAddRule.Name = "btnAddRule";
            this.btnAddRule.Width = 150;

            this.btnEditRule.Location = new System.Drawing.Point(176, 90);
            this.btnEditRule.Text = "Modifica";
            this.btnEditRule.Name = "btnEditRule";
            this.btnEditRule.Width = 100;

            this.btnDeleteRule.Location = new System.Drawing.Point(282, 90);
            this.btnDeleteRule.Text = "Elimina";
            this.btnDeleteRule.Name = "btnDeleteRule";
            this.btnDeleteRule.Width = 100;

            this.btnApplyPhase4.Location = new System.Drawing.Point(16, 130);
            this.btnApplyPhase4.Text = "Applica Fase 4";
            this.btnApplyPhase4.Name = "btnApplyPhase4";
            this.btnApplyPhase4.Width = 150;

            this.tabPhase4.Controls.Add(this.txtPhase4OutDir);
            this.tabPhase4.Controls.Add(this.btnBrowsePhase4Out);
            this.tabPhase4.Controls.Add(this.btnOpenPhase4Folder);
            this.tabPhase4.Controls.Add(this.btnReloadFase3Headers);
            this.tabPhase4.Controls.Add(this.btnAddRule);
            this.tabPhase4.Controls.Add(this.btnEditRule);
            this.tabPhase4.Controls.Add(this.btnDeleteRule);
            this.tabPhase4.Controls.Add(this.btnApplyPhase4);

            // ===== add to form
            this.Controls.Add(this.tabs);
            this.Controls.Add(this.lstLog);

            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }

