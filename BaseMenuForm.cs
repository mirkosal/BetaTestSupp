using System;
using System.Linq;
using System.Windows.Forms;

public partial class BaseMenuForm : Form
{
    // Opacità uniforme per tutti i form
    private const double DefaultOpacity = 0.95;

    // Content host esposto ai form ereditati (se volessero aggiungere controlli direttamente)
    protected Panel ContentHost => contentHost;

    public BaseMenuForm()
    {
        InitializeComponent();

        // Opacità comune
        try { this.Opacity = DefaultOpacity; } catch { /* alcuni ambienti potrebbero non supportare */ }

        // Quando l'handle è pronto, sposta i controlli esistenti dentro ContentHost (eccetto menu e host stesso)
        this.HandleCreated += (_, __) => MoveExistingControlsIntoContentHost();

        // In casi particolari (designer/custom init) ripetiamo allo Shown
        this.Shown += (_, __) => MoveExistingControlsIntoContentHost();
    }

    /// <summary>
    /// Sposta tutti i controlli già presenti nel form (eccetto MenuStrip e ContentHost)
    /// dentro il pannello ContentHost, così stanno sempre sotto il menu.
    /// </summary>
    private void MoveExistingControlsIntoContentHost()
    {
        // Se non è ancora stato fatto, reparent dei figli
        var toMove = this.Controls
            .Cast<Control>()
            .Where(c => c != menu && c != contentHost)
            .ToList();

        if (toMove.Count == 0) return;

        // Sospende il layout per evitare flicker
        this.SuspendLayout();
        contentHost.SuspendLayout();

        foreach (var ctrl in toMove)
        {
            ctrl.Parent = contentHost;
        }

        contentHost.ResumeLayout(performLayout: true);
        this.ResumeLayout(performLayout: true);
    }

    /// <summary>
    /// Mostra il nuovo form e chiude l'attuale in modo sicuro (una finestra alla volta).
    /// </summary>
    /// <summary>
    /// Passa ad un altro form usando il contesto globale dell'app.
    /// </summary>
    protected void SwitchTo(Form next)
    {
        Program.CurrentContext.SwitchTo(next);
    }

    // === Voci di menu comuni ===
    protected virtual void OnOpenCourseCompare(object? sender, EventArgs e)
    {
        // Se già aperto, passa a quello
        foreach (Form f in Application.OpenForms)
        {
            if (f is CourseCompareForm existing)
            {
                existing.Activate();
                SwitchTo(existing);
                return;
            }
        }
        // Altrimenti crea e apri
        var frm = new CourseCompareForm();
        SwitchTo(frm);
    }
    protected virtual void OnOpenCfCourseGenerator(object? sender, EventArgs e)
    {
        foreach (Form f in Application.OpenForms)
        {
            if (f is CfCourseGeneratorForm existing)
            {
                existing.Activate();
                Program.CurrentContext.SwitchTo(existing);
                return;
            }
        }

        var frm = new CfCourseGeneratorForm();
        Program.CurrentContext.SwitchTo(frm);
    }

    protected virtual void OnOpenCsvConverter(object? sender, EventArgs e)
    {
        // Se già aperto, passa direttamente a quello
        foreach (Form f in Application.OpenForms)
            if (f is CsvToXlsxForm existing) { Program.CurrentContext.SwitchTo(existing); return; }

        Program.CurrentContext.SwitchTo(new CsvToXlsxForm());
    }

    protected virtual void OnBackToMain(object? sender, EventArgs e)
    {
        foreach (Form f in Application.OpenForms)
            if (f is MainForm existing) { Program.CurrentContext.SwitchTo(existing); return; }

        Program.CurrentContext.SwitchTo(new MainForm());
    }

    protected virtual void OnExit(object? sender, EventArgs e)
    {
        Application.Exit();
    }
}
