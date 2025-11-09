// File: Core/Help.RuleRepository.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace BetaTestSupp.Core
{
    public sealed class HelpRuleRepository : IRuleRepository
    {
        private readonly string _file;

        public HelpRuleRepository(string appName, string fileName)
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName);
            Directory.CreateDirectory(dir);
            _file = Path.Combine(dir, fileName);
        }

        public IEnumerable<RuleDef> List()
        {
            try
            {
                if (!File.Exists(_file)) return Enumerable.Empty<RuleDef>();
                var json = File.ReadAllText(_file, new UTF8Encoding(false));
                var list = JsonSerializer.Deserialize<List<RuleDef>>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new List<RuleDef>();
                return list;
            }
            catch { return Enumerable.Empty<RuleDef>(); }
        }

        public void Save(RuleDef rule)
        {
            var all = List().ToList();
            var existing = all.FirstOrDefault(r => r.Id.Equals(rule.Id, StringComparison.OrdinalIgnoreCase));
            if (existing == null) all.Add(rule);
            else
            {
                existing.Name = rule.Name;
                existing.Kind = rule.Kind;
                existing.Field1 = rule.Field1;
                existing.Field2 = rule.Field2;
                existing.IntParam = rule.IntParam;
                existing.Phase2Dataset = rule.Phase2Dataset;
            }
            Write(all);
        }

        public void Delete(string id)
        {
            var all = List().ToList();
            all.RemoveAll(r => r.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
            Write(all);
        }

        private void Write(List<RuleDef> rules)
        {
            var json = JsonSerializer.Serialize(rules, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_file, json, new UTF8Encoding(false));
        }
    }
}
