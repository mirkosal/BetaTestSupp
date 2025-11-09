using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BetaTestSupp.Core;

public partial class RuleEditorForm : Form
{
    private readonly List<string> _availableFields;
    private RuleDef _rule;

    public RuleDef Result => _rule;

    public RuleEditorForm(List<string> availableFields, RuleDef? ruleToEdit = null)
    {
        _availableFields = availableFields ?? new();
        _rule = ruleToEdit != null ? Clone(ruleToEdit) : new RuleDef();

        InitializeComponent();

        Load += (_, __) =>
        {
            lblTitle.Text = ruleToEdit == null ? "Nuova regola" : "Modifica regola";

            // Tipi
            cmbKind.Items.Clear();
            cmbKind.Items.AddRange(Enum.GetNames(typeof(RuleKind)));
            var kindName = Enum.GetName(typeof(RuleKind), _rule.Kind) ?? nameof(RuleKind.DateAfterToday);
            cmbKind.SelectedItem = kindName;

            // Campi (da headers Fase 3)
            void fill(ComboBox cmb)
            {
                cmb.Items.Clear();
                foreach (var h in _availableFields) cmb.Items.Add(h);
                if (cmb.Items.Count > 0) cmb.SelectedIndex = 0;
            }
            fill(cmbField1);
            fill(cmbField2);

            txtName.Text = _rule.Name ?? "";
            cmbDataset.SelectedItem = _rule.Phase2Dataset ?? "";
            if (!string.IsNullOrWhiteSpace(_rule.Field1)) cmbField1.SelectedItem = _rule.Field1;
            if (!string.IsNullOrWhiteSpace(_rule.Field2)) cmbField2.SelectedItem = _rule.Field2;
            if (_rule.IntParam.HasValue) numIntParam.Value = Math.Max(numIntParam.Minimum, Math.Min(numIntParam.Maximum, _rule.IntParam.Value));

            UpdateVisibility();
        };

        cmbKind.SelectedIndexChanged += (_, __) => UpdateVisibility();
        btnOk.Click += (_, __) => Confirm();
    }

    private static RuleDef Clone(RuleDef r) => new RuleDef
    {
        Id = r.Id,
        Name = r.Name,
        Kind = r.Kind,
        Field1 = r.Field1,
        Field2 = r.Field2,
        IntParam = r.IntParam,
        Phase2Dataset = r.Phase2Dataset
    };

    private void UpdateVisibility()
    {
        var kind = (cmbKind.SelectedItem?.ToString() ?? nameof(RuleKind.DateAfterToday));
        bool needsDataset = kind is nameof(RuleKind.NotPresentInPhase2) or nameof(RuleKind.PairNotPresentInPhase2);
        bool needsField2 = kind is nameof(RuleKind.PairNotPresentInPhase2);
        bool needsInt = kind is nameof(RuleKind.MaxLength);

        lblDataset.Enabled = cmbDataset.Enabled = needsDataset;
        lblField2.Enabled = cmbField2.Enabled = needsField2;
        lblIntParam.Enabled = numIntParam.Enabled = needsInt;
    }

    private void Confirm()
    {
        string name = (txtName.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(this, "Inserisci un nome regola.", "Attenzione");
            return;
        }
        var kind = (cmbKind.SelectedItem?.ToString() ?? nameof(RuleKind.DateAfterToday));
        var parsed = Enum.TryParse<RuleKind>(kind, out var k) ? k : RuleKind.DateAfterToday;

        _rule.Name = name;
        _rule.Kind = parsed;
        _rule.Field1 = cmbField1.SelectedItem?.ToString() ?? "";
        _rule.Field2 = cmbField2.SelectedItem?.ToString();
        _rule.IntParam = (int)numIntParam.Value;
        _rule.Phase2Dataset = cmbDataset.SelectedItem?.ToString();

        DialogResult = DialogResult.OK;
        Close();
    }
}
