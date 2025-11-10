// RuleEditorForm.cs — Form editor regole (SENZA namespace)
// - Integra con IRuleRepository via riflessione (metodi tipici GetAll/Save/Delete)
// - Fallback JSON interno se repo assente (stesso formato di CfCourseGeneratorForm)

#nullable enable
using BetaTestSupp.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

public partial class RuleEditorForm : Form
{
    private readonly IRuleRepository? _repo;
    private readonly string _storePath;
    public bool DeleteMode { get; set; } = false;

    public RuleEditorForm(IRuleRepository? repo, string? preselectName)
    {
        InitializeComponent();
        _repo = repo;

        var appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BetaTestSupp");
        Directory.CreateDirectory(appData);
        _storePath = Path.Combine(appData, "CfCourseGenerator.rules.json");

        this.Load += (_, __) =>
        {
            LoadRulesIntoList(preselectName);
            if (DeleteMode)
            {
                grpEdit.Enabled = false;
                btnSave.Enabled = false;
                Text = "Elimina regola";
            }
        };

        lstRules.SelectedIndexChanged += (_, __) =>
        {
            if (DeleteMode) return;
            var m = lstRules.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(m)) return;
            var model = GetAll().FirstOrDefault(r => string.Equals(r.Name, m, StringComparison.OrdinalIgnoreCase));
            if (model != null)
            {
                txtName.Text = model.Name;
                txtType.Text = model.Type;
                txtExpr.Text = model.Expression;
            }
        };

        btnNew.Click += (_, __) =>
        {
            lstRules.ClearSelected();
            txtName.Text = "";
            txtType.Text = "";
            txtExpr.Text = "";
            txtName.Focus();
        };

        btnSave.Click += (_, __) =>
        {
            var model = new RuleModelFlexible
            {
                Name = txtName.Text.Trim(),
                Type = txtType.Text.Trim(),
                Expression = txtExpr.Text.Trim()
            };
            if (string.IsNullOrWhiteSpace(model.Name))
            {
                MessageBox.Show(this, "Inserire il Nome regola.", "Dati mancanti", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Save(model))
            {
                MessageBox.Show(this, "Salvataggio non riuscito (repo assente?). Usato fallback JSON.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            LoadRulesIntoList(model.Name);
        };

        btnDelete.Click += (_, __) =>
        {
            var sel = lstRules.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(sel))
            {
                MessageBox.Show(this, "Seleziona una regola.", "Elimina", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(this, $"Eliminare '{sel}'?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (!DeleteByName(sel))
                    MessageBox.Show(this, "Eliminazione non riuscita su repo; usato fallback JSON.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                LoadRulesIntoList(null);
            }
        };

        btnClose.Click += (_, __) => this.DialogResult = DialogResult.OK;
    }

    private void LoadRulesIntoList(string? toSelect)
    {
        var rules = GetAll().OrderBy(r => r.Name).ToList();

        lstRules.BeginUpdate();
        lstRules.Items.Clear();
        foreach (var r in rules) lstRules.Items.Add(r.Name);
        lstRules.EndUpdate();

        if (!string.IsNullOrWhiteSpace(toSelect))
        {
            int idx = lstRules.Items.IndexOf(toSelect);
            if (idx >= 0) lstRules.SelectedIndex = idx;
        }
    }

    private List<RuleModelFlexible> GetAll()
    {
        var list = new List<RuleModelFlexible>();
        if (_repo != null)
        {
            try
            {
                var t = _repo.GetType();
                var m = t.GetMethod("GetAllRules", Type.EmptyTypes) ??
                        t.GetMethod("GetAll", Type.EmptyTypes) ??
                        t.GetMethod("List", Type.EmptyTypes);
                if (m != null)
                {
                    var res = m.Invoke(_repo, null) as System.Collections.IEnumerable;
                    if (res != null)
                    {
                        foreach (var o in res)
                        {
                            var r = ToFlexible(o);
                            if (r != null) list.Add(r);
                        }
                        return list;
                    }
                }
            }
            catch { /* fallback json */ }
        }

        // Fallback JSON
        try
        {
            if (!File.Exists(_storePath)) return list;
            var json = File.ReadAllText(_storePath);
            var arr = JsonSerializer.Deserialize<RuleModelFlexible[]>(json) ?? Array.Empty<RuleModelFlexible>();
            list.AddRange(arr);
        }
        catch { }
        return list;
    }

    private static RuleModelFlexible? ToFlexible(object? obj)
    {
        if (obj == null) return null;
        var t = obj.GetType();
        string name = t.GetProperty("Name")?.GetValue(obj)?.ToString()
                      ?? t.GetProperty("Title")?.GetValue(obj)?.ToString() ?? "Rule";
        string type = t.GetProperty("Type")?.GetValue(obj)?.ToString() ?? "";
        string expr = t.GetProperty("Expression")?.GetValue(obj)?.ToString() ?? "";
        return new RuleModelFlexible { Name = name, Type = type, Expression = expr };
    }

    private bool Save(RuleModelFlexible model)
    {
        // 1) Prova repository
        if (_repo != null)
        {
            try
            {
                var t = _repo.GetType();
                var m = t.GetMethod("SaveRule", new[] { model.GetType() }) ??
                        t.GetMethod("Upsert", new[] { model.GetType() });
                if (m != null) { m.Invoke(_repo, new object[] { model }); return true; }
            }
            catch { /* fallback json */ }
        }

        // 2) Fallback JSON
        try
        {
            var list = GetAll();
            var existing = list.FirstOrDefault(x => string.Equals(x.Name, model.Name, StringComparison.OrdinalIgnoreCase));
            if (existing == null) list.Add(model);
            else { existing.Type = model.Type; existing.Expression = model.Expression; }

            var json = JsonSerializer.Serialize(list.ToArray(), new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_storePath, json);
            return false; // false = non via repo, ma comunque salvato
        }
        catch { return false; }
    }

    private bool DeleteByName(string name)
    {
        // 1) Prova repository
        if (_repo != null)
        {
            try
            {
                var t = _repo.GetType();
                var m = t.GetMethod("DeleteRuleByName", new[] { typeof(string) }) ??
                        t.GetMethod("DeleteByName", new[] { typeof(string) }) ??
                        t.GetMethod("Delete", new[] { typeof(string) });
                if (m != null) { m.Invoke(_repo, new object[] { name }); return true; }
            }
            catch { /* fallback json */ }
        }

        // 2) Fallback JSON
        try
        {
            var list = GetAll();
            list.RemoveAll(r => string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase));
            var json = JsonSerializer.Serialize(list.ToArray(), new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_storePath, json);
            return false;
        }
        catch { return false; }
    }
}
