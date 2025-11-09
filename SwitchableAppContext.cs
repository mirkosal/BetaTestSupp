using System;
using System.Windows.Forms;

/// <summary>
/// ApplicationContext che permette di cambiare il "form principale"
/// senza terminare l'applicazione.
/// </summary>
public sealed class SwitchableAppContext : ApplicationContext
{
    public SwitchableAppContext()
    {
    }

    /// <summary>
    /// Passa a un nuovo form: imposta il MainForm sul nuovo, poi chiude e dispose il precedente.
    /// </summary>
    public void SwitchTo(Form next)
    {
        if (next == null) throw new ArgumentNullException(nameof(next));

        next.StartPosition = FormStartPosition.CenterScreen;

        // Vecchio MainForm
        var old = this.MainForm;

        // Scollega l'handler dal precedente (se c'era)
        if (old != null)
            old.FormClosed -= OnFormClosed;

        // Imposta e mostra il nuovo come MainForm
        this.MainForm = next;
        this.MainForm.FormClosed += OnFormClosed;
        this.MainForm.Show();
        this.MainForm.Activate();

        // Chiudi/Dispose il precedente DOPO aver cambiato MainForm
        if (old != null)
        {
            try
            {
                // chiusura asincrona per evitare race nel message loop
                old.BeginInvoke(new Action(() =>
                {
                    try { old.Close(); } catch { }
                    try { old.Dispose(); } catch { }
                }));
            }
            catch { /* ignore */ }
        }
    }

    private void OnFormClosed(object? sender, FormClosedEventArgs e)
    {

    }
}
