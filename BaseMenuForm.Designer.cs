using System.Drawing;
using System.Windows.Forms;

public partial class BaseMenuForm : Form
{
    private MenuStrip menu;
    private ToolStripMenuItem mFile;
    private ToolStripMenuItem mExit;
    private ToolStripMenuItem mOps;
    private ToolStripMenuItem mCsvToXlsx;
    private ToolStripMenuItem mCourseCompare;
    private ToolStripMenuItem mCfCourseGen;
    private ToolStripMenuItem mNav;
    private ToolStripMenuItem mBackToMain;
    private Panel contentHost;

    private void InitializeComponent()
    {
        menu = new MenuStrip();
        mFile = new ToolStripMenuItem("&File");
        mExit = new ToolStripMenuItem("Esci");
        mOps = new ToolStripMenuItem("&Operazioni");
        mCsvToXlsx = new ToolStripMenuItem("Converti CSV→XLSX…");
        mCourseCompare = new ToolStripMenuItem("Valutazione corsi…");
        mCfCourseGen = new ToolStripMenuItem("Genera CF/Corsi…");   // ✅ nuova voce
        mNav = new ToolStripMenuItem("&Navigazione");
        mBackToMain = new ToolStripMenuItem("Torna al Main");
        contentHost = new Panel();

        SuspendLayout();

        // === MenuStrip ===
        menu.ImageScalingSize = new Size(20, 20);
        menu.Items.AddRange(new ToolStripItem[] { mFile, mOps, mNav });
        menu.Dock = DockStyle.Top;
        menu.BackColor = Color.FromArgb(30, 30, 36);
        menu.ForeColor = Color.White;

        // === FILE ===
        mFile.DropDownItems.AddRange(new ToolStripItem[] { mExit });
        mExit.Click += OnExit;

        // === OPERAZIONI ===
        mOps.DropDownItems.AddRange(new ToolStripItem[]
        {
            mCsvToXlsx,
            mCourseCompare,
            mCfCourseGen // ✅ aggiunto qui
        });

        mCsvToXlsx.Click += OnOpenCsvConverter;
        mCourseCompare.Click += OnOpenCourseCompare;
        mCfCourseGen.Click += OnOpenCfCourseGenerator; // ✅ evento click

        // === NAVIGAZIONE ===
        mNav.DropDownItems.AddRange(new ToolStripItem[] { mBackToMain });
        mBackToMain.Click += OnBackToMain;

        // === Content Host ===
        contentHost.Dock = DockStyle.Fill;
        contentHost.BackColor = Color.FromArgb(24, 24, 28);
        contentHost.ForeColor = Color.Lime;
        contentHost.Padding = new Padding(12);

        // === Base Form ===
        AutoScaleMode = AutoScaleMode.None;
        BackColor = Color.FromArgb(24, 24, 28);
        ForeColor = Color.Lime;
        Font = new Font("Segoe UI", 9F);
        Controls.Add(contentHost);
        Controls.Add(menu);
        MainMenuStrip = menu;
        Text = "Applicazione Gestione File";
        MinimumSize = new Size(900, 600);

        ResumeLayout(false);
        PerformLayout();
    }
}
