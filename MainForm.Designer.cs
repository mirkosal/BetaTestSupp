using System.Windows.Forms;
using System.Xml.Linq;


    public partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        // Root
        private SplitContainer splitRoot;

        // Top bar
        private Panel pTopContainer;
        private FlowLayoutPanel pTop;
        internal Button btnOpen;
        internal Button btnRefresh;
        private Label lblSheetCaption;
        internal ComboBox cmbSheets;
        private Label lblColCaption;
        internal ComboBox cmbDisplayColumn;
        private Label lblRowsCaption;
        internal ComboBox cmbRows;
        private Label lblScreenQuickCaption;
        internal ComboBox cmbScreenQuick;
        internal Button btnScreenQuickSet;
        internal CheckBox chkHideOnCapture;

        // Center/bottom
        private TableLayoutPanel tlpMain;
        internal Label lblInfo;
        private Panel pValueWrap;
        internal Label lblBValue;
        internal ListBox lstKeyValues;

        // Bottom panel: add/move column
        private TableLayoutPanel pBottom;
        private Label lblNewCol;
        internal TextBox txtNewColName;
        private Label lblRel;
        internal ComboBox cmbNewColRel;
        private Label lblAnchor;
        internal ComboBox cmbNewColAnchor;
        private Label lblColW;
        internal NumericUpDown nudNewColWidth;
        private Label lblRowH;
        internal NumericUpDown nudNewRowHeight;
        internal CheckBox chkMoveIfExists;
        internal Button btnAddCol;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            // ===== Form =====
            SuspendLayout();
            Name = "MainForm";
            Text = "Overlay Excel — ALT+S scatta, ALT+N riga+1, ALT+M riga-1";
            StartPosition = FormStartPosition.Manual;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            Width = 900;
            Height = 650;
            MinimumSize = new System.Drawing.Size(760, 520);
            Opacity = 0.9;
            BackColor = System.Drawing.Color.FromArgb(24, 24, 28);
            ForeColor = System.Drawing.Color.Lime;
            AutoScaleMode = AutoScaleMode.None;
            AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            DoubleBuffered = true;

            // ===== Split root =====
            splitRoot = new SplitContainer
            {
                Name = "splitRoot",
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                FixedPanel = FixedPanel.Panel1,
                IsSplitterFixed = false,
                SplitterWidth = 4,
                Panel1MinSize = 210,
                BackColor = System.Drawing.Color.FromArgb(24, 24, 28),
                BorderStyle = BorderStyle.None
            };
            splitRoot.SuspendLayout();
            Controls.Add(splitRoot);
            splitRoot.SplitterDistance = 210;

            // ===== Top container =====
            pTopContainer = new Panel
            {
                Name = "pTopContainer",
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                Margin = new Padding(0)
            };
            pTopContainer.SuspendLayout();
            splitRoot.Panel1.Controls.Add(pTopContainer);

            // ===== Top FlowLayout (toolbar) =====
            pTop = new FlowLayoutPanel
            {
                Name = "pTop",
                Dock = DockStyle.Fill,
                Padding = new Padding(8),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,        // va a capo su più righe (niente scroll)
                AutoScroll = false,
                AutoSize = false,
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                Margin = new Padding(0)
            };
            pTop.SuspendLayout();
            pTopContainer.Controls.Add(pTop);

            btnOpen = MakeBtn("Apri Excel…", 120, System.Drawing.Color.Orange, System.Drawing.Color.Black);
            btnOpen.Name = "btnOpen";
            btnOpen.TabIndex = 0;
            pTop.Controls.Add(btnOpen);

            btnRefresh = MakeBtn("Aggiorna", 100);
            btnRefresh.Name = "btnRefresh";
            btnRefresh.TabIndex = 1;
            pTop.Controls.Add(btnRefresh);

            lblSheetCaption = MakeLblTop("Foglio:");
            lblSheetCaption.Name = "lblSheetCaption";
            pTop.Controls.Add(lblSheetCaption);

            cmbSheets = MakeCmb(200);
            cmbSheets.Name = "cmbSheets";
            cmbSheets.TabIndex = 2;
            pTop.Controls.Add(cmbSheets);

            lblColCaption = MakeLblTop("Colonna vis.:");
            lblColCaption.Name = "lblColCaption";
            pTop.Controls.Add(lblColCaption);

            cmbDisplayColumn = MakeCmb(220);
            cmbDisplayColumn.Name = "cmbDisplayColumn";
            cmbDisplayColumn.TabIndex = 3;
            pTop.Controls.Add(cmbDisplayColumn);

            lblRowsCaption = MakeLblTop("Righe:");
            lblRowsCaption.Name = "lblRowsCaption";
            pTop.Controls.Add(lblRowsCaption);

            cmbRows = MakeCmb(160);
            cmbRows.Name = "cmbRows";
            cmbRows.TabIndex = 4;
            pTop.Controls.Add(cmbRows);

            lblScreenQuickCaption = MakeLblTop("Screenshot in:");
            lblScreenQuickCaption.Name = "lblScreenQuickCaption";
            pTop.Controls.Add(lblScreenQuickCaption);

            cmbScreenQuick = MakeCmb(220);
            cmbScreenQuick.Name = "cmbScreenQuick";
            cmbScreenQuick.TabIndex = 5;
            pTop.Controls.Add(cmbScreenQuick);

            btnScreenQuickSet = MakeBtn("Imposta", 100);
            btnScreenQuickSet.Name = "btnScreenQuickSet";
            btnScreenQuickSet.TabIndex = 6;
            pTop.Controls.Add(btnScreenQuickSet);

            chkHideOnCapture = new CheckBox
            {
                Name = "chkHideOnCapture",
                Text = "Nascondi overlay durante scatto",
                AutoSize = true,
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(8, 6, 0, 0),
                TabIndex = 7
            };
            pTop.Controls.Add(chkHideOnCapture);

            pTop.ResumeLayout(false);
            pTop.PerformLayout();
            pTopContainer.ResumeLayout(false);

            // ===== Panel2 (info + label + list + bottom) =====
            tlpMain = new TableLayoutPanel
            {
                Name = "tlpMain",
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = System.Drawing.Color.FromArgb(24, 24, 28),
                Margin = new Padding(0)
            };
            tlpMain.SuspendLayout();
            splitRoot.Panel2.Controls.Add(tlpMain);

            tlpMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F)); // info
            tlpMain.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // label auto
            tlpMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // list fill
            tlpMain.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // bottom auto

            // info
            lblInfo = new Label
            {
                Name = "lblInfo",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Segoe UI", 9F),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Text = "—",
                ForeColor = System.Drawing.Color.Silver,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(0),
                TabIndex = 10
            };
            tlpMain.Controls.Add(lblInfo, 0, 0);

            // wrapper label valore
            pValueWrap = new Panel
            {
                Name = "pValueWrap",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                Margin = new Padding(0)
            };
            pValueWrap.SuspendLayout();
            tlpMain.Controls.Add(pValueWrap, 0, 1);

            lblBValue = new Label
            {
                Name = "lblBValue",
                AutoSize = true,
                Dock = DockStyle.Top,
                Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Text = "—",
                ForeColor = System.Drawing.Color.Yellow,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(8),
                UseCompatibleTextRendering = true,
                TabIndex = 11
            };
            pValueWrap.Controls.Add(lblBValue);
            pValueWrap.ResumeLayout(false);
            pValueWrap.PerformLayout();

            // lista chiave/valore
            lstKeyValues = new ListBox
            {
                Name = "lstKeyValues",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Consolas", 9.5F),
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,             // evita tagli
                HorizontalScrollbar = false,
                Margin = new Padding(8, 6, 8, 6),
                TabIndex = 12
            };
            tlpMain.Controls.Add(lstKeyValues, 0, 2);

            // ===== Bottom panel (auto-size) =====
            pBottom = new TableLayoutPanel
            {
                Name = "pBottom",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 12,
                RowCount = 2,
                Padding = new Padding(8),
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                Margin = new Padding(0, 0, 0, 8)
            };
            pBottom.SuspendLayout();
            tlpMain.Controls.Add(pBottom, 0, 3);

            // colonne
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // lbl new
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240)); // txt
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // lbl rel
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140)); // cmb rel
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // lbl anchor
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260)); // cmb anchor
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // lbl colW
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));  // nud colW
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // lbl rowH
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));  // nud rowH
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // chk move
            pBottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // btn add
            pBottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            pBottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            lblNewCol = new Label
            {
                Name = "lblNewCol",
                Text = "Nuovo campo:",
                AutoSize = true,
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(0, 8, 6, 0)
            };
            pBottom.Controls.Add(lblNewCol, 0, 0);

            txtNewColName = new TextBox
            {
                Name = "txtNewColName",
                PlaceholderText = "es. screen / notes",
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                ForeColor = System.Drawing.Color.Lime,
                BorderStyle = BorderStyle.FixedSingle,
                Width = 240,
                Margin = new Padding(0, 4, 8, 0),
                TabIndex = 20
            };
            pBottom.Controls.Add(txtNewColName, 1, 0);

            lblRel = new Label
            {
                Name = "lblRel",
                Text = "Posizione:",
                AutoSize = true,
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(0, 8, 6, 0)
            };
            pBottom.Controls.Add(lblRel, 2, 0);

            cmbNewColRel = MakeCmb(140);
            cmbNewColRel.Name = "cmbNewColRel";
            cmbNewColRel.TabIndex = 21;
            pBottom.Controls.Add(cmbNewColRel, 3, 0);

            lblAnchor = new Label
            {
                Name = "lblAnchor",
                Text = "Colonna di riferimento:",
                AutoSize = true,
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(8, 8, 6, 0)
            };
            pBottom.Controls.Add(lblAnchor, 4, 0);

            cmbNewColAnchor = MakeCmb(260);
            cmbNewColAnchor.Name = "cmbNewColAnchor";
            cmbNewColAnchor.TabIndex = 22;
            pBottom.Controls.Add(cmbNewColAnchor, 5, 0);

            lblColW = new Label
            {
                Name = "lblColW",
                Text = "Larg. col.:",
                AutoSize = true,
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(8, 8, 4, 0)
            };
            pBottom.Controls.Add(lblColW, 6, 0);

            nudNewColWidth = new NumericUpDown
            {
                Name = "nudNewColWidth",
                Minimum = 0,
                Maximum = 255,
                DecimalPlaces = 1,
                Increment = 0.5M,
                Value = 0,
                Width = 90,
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                ForeColor = System.Drawing.Color.Lime,
                BorderStyle = BorderStyle.FixedSingle,
                TabIndex = 23
            };
            pBottom.Controls.Add(nudNewColWidth, 7, 0);

            lblRowH = new Label
            {
                Name = "lblRowH",
                Text = "Alt. righe (pt):",
                AutoSize = true,
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(8, 8, 4, 0)
            };
            pBottom.Controls.Add(lblRowH, 8, 0);

            nudNewRowHeight = new NumericUpDown
            {
                Name = "nudNewRowHeight",
                Minimum = 0,
                Maximum = 409,
                DecimalPlaces = 1,
                Increment = 0.5M,
                Value = 0,
                Width = 90,
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                ForeColor = System.Drawing.Color.Lime,
                BorderStyle = BorderStyle.FixedSingle,
                TabIndex = 24
            };
            pBottom.Controls.Add(nudNewRowHeight, 9, 0);

            chkMoveIfExists = new CheckBox
            {
                Name = "chkMoveIfExists",
                Text = "Se esiste, sposta",
                AutoSize = true,
                Checked = true,
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(8, 8, 8, 0),
                TabIndex = 25
            };
            pBottom.Controls.Add(chkMoveIfExists, 10, 0);

            btnAddCol = new Button
            {
                Name = "btnAddCol",
                Text = "Aggiungi/Imposta",
                AutoSize = true,
                BackColor = System.Drawing.Color.FromArgb(40, 40, 50),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Standard,
                Margin = new Padding(8, 4, 4, 4),
                Padding = new Padding(8, 2, 8, 2),
                TabIndex = 26
            };
            pBottom.Controls.Add(btnAddCol, 11, 0);

            pBottom.ResumeLayout(false);
            pBottom.PerformLayout();

            tlpMain.ResumeLayout(false);
            tlpMain.PerformLayout();

            splitRoot.ResumeLayout(false);

            ResumeLayout(false);
            PerformLayout();
        }

        // === Helpers grafici ===
        private static Button MakeBtn(string text, int width, System.Drawing.Color? back = null, System.Drawing.Color? fore = null)
        {
            return new Button
            {
                Text = text,
                AutoSize = false,
                Size = new System.Drawing.Size(width, 28),
                MinimumSize = new System.Drawing.Size(width, 28),
                MaximumSize = new System.Drawing.Size(width, 28),
                BackColor = back ?? System.Drawing.Color.FromArgb(40, 40, 50),
                ForeColor = fore ?? System.Drawing.Color.White,
                FlatStyle = FlatStyle.Standard,
                Margin = new Padding(0, 0, 8, 0),
                Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold),
                UseVisualStyleBackColor = true
            };
        }

        private static Label MakeLblTop(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Padding = new Padding(0, 6, 4, 0),
                ForeColor = System.Drawing.Color.Lime,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(0, 0, 8, 0)
            };
        }

        private static ComboBox MakeCmb(int width)
        {
            return new ComboBox
            {
                Width = width,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = System.Drawing.Color.FromArgb(20, 20, 24),
                ForeColor = System.Drawing.Color.Lime,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 8, 0),
                IntegralHeight = false,
                MaxDropDownItems = 20
            };
        }

        #endregion
    }
