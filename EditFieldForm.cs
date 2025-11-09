using System;
using System.Windows.Forms;

public partial class EditFieldForm : Form
{
    /// <summary>Nuovo testo immesso (ritornato se DialogResult == OK).</summary>
    public string? NewText => txtValue.Text;

    private readonly string _contextTitle;
    private readonly string _fieldDesc;
    private readonly string _oldName;

    /// <summary>
    /// Costruttore generale: imposta titolo, vecchio nome (a sinistra) e
    /// precompila la TextBox con il valore corrente.
    /// </summary>
    public EditFieldForm(string contextTitle, string fieldDescription, string? currentText)
    {
        _contextTitle = contextTitle ?? "Fase";
        _fieldDesc = fieldDescription ?? "Campo";
        _oldName = currentText ?? "";

        InitializeComponent();

        Load += (_, __) =>
        {
            // Titolo in alto come richiesto: "Rinomina campo — <fieldDescription>"
            Text = "Rinomina campo";
            lblTitle.Text = $"Rinomina campo — {_fieldDesc}";
            lblOldName.Text = string.IsNullOrWhiteSpace(_oldName) ? "(senza nome)" : _oldName;

            // Precompila con il nome corrente
            txtValue.Text = _oldName;
            txtValue.SelectAll();
            txtValue.Focus();
        };

        // Click bottoni
        btnOk.Click += (_, __) => Confirm();
        btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };

        // Invio = conferma
        txtValue.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter && !e.Shift && !e.Control && !e.Alt)
            {
                e.SuppressKeyPress = true;
                Confirm();
            }
        };
    }

    /// <summary>
    /// Overload compatibile con la firma precedente:
    /// new EditFieldForm("Fase3", "Nome campo (intestazione)", tf.DisplayName, true)
    /// Il bool finale è ignorato: serve solo a distinguere la signature.
    /// </summary>
    public EditFieldForm(string contextTitle, string fieldDescription, string currentText, bool _)
        : this(contextTitle, fieldDescription, currentText)
    {
    }

    private void Confirm()
    {
        var name = (txtValue.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(this, "Inserisci un nome valido.", "Attenzione",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtValue.Focus();
            return;
        }
        DialogResult = DialogResult.OK;
        Close();
    }
}
