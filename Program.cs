using System;
using System.Windows.Forms;

internal static class Program
{
    // Context globale per poter fare lo switch dei form senza chiudere l'app
    public static SwitchableAppContext CurrentContext = null!;

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        CurrentContext = new SwitchableAppContext();

        // Avvia direttamente col MainForm
        CurrentContext.SwitchTo(new MainForm());

        Application.Run(CurrentContext);
    }
}
